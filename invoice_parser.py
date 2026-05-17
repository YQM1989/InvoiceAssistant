from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pypdfium2 as pdfium
from pypdf import PdfReader
from rapidocr_onnxruntime import RapidOCR

from storage import clean_company_name, is_valid_company_name, normalize_company, normalize_invoice_number


@dataclass
class ParsedInvoice:
    seller_name: str
    buyer_name: str
    invoice_number: str
    amount: float
    invoice_date: str
    raw_text: str
    seller_in_whitelist: bool
    buyer_in_whitelist: bool
    ocr_used: bool = False
    issues: list[str] = field(default_factory=list)


@dataclass
class OCRToken:
    text: str
    score: float
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


INVALID_FILENAME_CHARS = '<>:"/\\|?*'
_OCR_ENGINE: RapidOCR | None = None


def sanitize_filename(name: str) -> str:
    cleaned = "".join("_" if ch in INVALID_FILENAME_CHARS else ch for ch in name)
    cleaned = cleaned.strip().rstrip(".")
    return cleaned or "未命名发票.pdf"


def read_pdf_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    chunks: list[str] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        if text:
            chunks.append(text)
    return "\n".join(chunks)


def _get_ocr_engine() -> RapidOCR:
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _render_pdf_page(pdf_path: Path, page_index: int = 0, scale: float = 3.0) -> np.ndarray:
    pdf = pdfium.PdfDocument(str(pdf_path))
    page = pdf[page_index]
    bitmap = page.render(scale=scale)
    if hasattr(bitmap, "to_numpy"):
        return bitmap.to_numpy()
    return np.array(bitmap.to_pil())


def read_pdf_ocr_tokens(pdf_path: Path, page_index: int = 0, scale: float = 3.0) -> tuple[list[OCRToken], str]:
    img = _render_pdf_page(pdf_path, page_index=page_index, scale=scale)
    engine = _get_ocr_engine()
    result, _ = engine(img)
    if not result:
        return [], ""

    h, w = img.shape[:2]
    tokens: list[OCRToken] = []
    for row in result:
        if len(row) < 3:
            continue
        box, text, score = row[0], str(row[1]).strip(), float(row[2] or 0.0)
        if not text:
            continue
        pts = np.array(box, dtype=float)
        x0 = max(0.0, float(pts[:, 0].min()) / max(1, w))
        y0 = max(0.0, float(pts[:, 1].min()) / max(1, h))
        x1 = min(1.0, float(pts[:, 0].max()) / max(1, w))
        y1 = min(1.0, float(pts[:, 1].max()) / max(1, h))
        tokens.append(OCRToken(text=text, score=score, x0=x0, y0=y0, x1=x1, y1=y1))

    tokens.sort(key=lambda token: (token.cy, token.cx))
    return tokens, "\n".join(token.text for token in tokens)


def normalize_date(raw: str) -> str:
    value = raw.strip()
    value = value.replace("年", "-").replace("月", "-").replace("日", "")
    value = value.replace("/", "-").replace(".", "-")
    parts = [p for p in value.split("-") if p]
    if len(parts) >= 3:
        y = parts[0].zfill(4)
        m = parts[1].zfill(2)
        d = parts[2].zfill(2)
        return f"{y}-{m}-{d}"
    return ""


def _compact_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "")


def _find_whitelist_in_text(text: str, whitelist_map: dict[str, str]) -> str:
    compact = normalize_company(text)
    if not compact:
        return ""
    matches: list[tuple[int, str]] = []
    for norm, standard_name in whitelist_map.items():
        if norm and norm in compact:
            matches.append((len(norm), standard_name))
    if not matches:
        return ""
    matches.sort(key=lambda item: item[0], reverse=True)
    return matches[0][1]


def _group_tokens_into_lines(tokens: list[OCRToken], y_threshold: float = 0.018) -> list[list[OCRToken]]:
    lines: list[list[OCRToken]] = []
    for token in sorted(tokens, key=lambda item: (item.cy, item.cx)):
        if not lines or abs(token.cy - np.mean([t.cy for t in lines[-1]])) > y_threshold:
            lines.append([token])
        else:
            lines[-1].append(token)
    for line in lines:
        line.sort(key=lambda item: item.cx)
    return lines


def _join_lines(tokens: list[OCRToken]) -> str:
    lines = _group_tokens_into_lines(tokens)
    return "\n".join(" ".join(token.text for token in line) for line in lines)


def _tokens_in_region(tokens: list[OCRToken], x0: float, y0: float, x1: float, y1: float) -> list[OCRToken]:
    return [
        token
        for token in tokens
        if token.cx >= x0 and token.cx <= x1 and token.cy >= y0 and token.cy <= y1
    ]


def _region_average_score(tokens: list[OCRToken]) -> float:
    if not tokens:
        return 0.0
    return sum(token.score for token in tokens) / len(tokens)


def _extract_company_from_region_text(region_text: str) -> str:
    patterns = [
        r"名称[:：]?\s*([^\n]{4,80}?)(?:统一社会信用代码|纳税人识别号|地址|电话|开户行|账号|项目名称|货物)",
        r"名称[:：]?\s*([^\n]{4,80})",
    ]
    compact_text = _compact_text(region_text)
    for pattern in patterns:
        match = re.search(pattern, compact_text, re.IGNORECASE)
        if match:
            candidate = clean_company_name(match.group(1))
            if is_valid_company_name(candidate):
                return candidate
    for line in region_text.splitlines():
        candidate = clean_company_name(line)
        if is_valid_company_name(candidate) and any(
            token in candidate for token in ["公司", "科技", "电子", "银行", "集团", "股份", "信息", "材料", "模塑", "制造"]
        ):
            return candidate
    return ""


def _extract_parties_from_fixed_layout(
    tokens: list[OCRToken],
    seller_whitelist_map: dict[str, str],
    buyer_whitelist_map: dict[str, str],
) -> tuple[str, bool, str, bool, list[str], str]:
    issues: list[str] = []
    if not tokens:
        return "", False, "", False, ["主体定位低置信度"], ""

    buyer_tokens = _tokens_in_region(tokens, 0.04, 0.14, 0.50, 0.40)
    seller_tokens = _tokens_in_region(tokens, 0.50, 0.14, 0.96, 0.40)

    buyer_text = _join_lines(buyer_tokens)
    seller_text = _join_lines(seller_tokens)
    layout_dump = f"[BUYER_REGION]\n{buyer_text}\n[SELLER_REGION]\n{seller_text}"

    if len(buyer_tokens) < 3 or len(seller_tokens) < 3:
        issues.append("版式不匹配")
    if _region_average_score(buyer_tokens) < 0.55 or _region_average_score(seller_tokens) < 0.55:
        issues.append("主体定位低置信度")

    buyer_name = _find_whitelist_in_text(buyer_text, buyer_whitelist_map)
    buyer_hit = bool(buyer_name)
    if not buyer_name:
        buyer_name = _extract_company_from_region_text(buyer_text)
        buyer_hit = normalize_company(buyer_name) in buyer_whitelist_map if buyer_name else False
        if not buyer_name:
            issues.append("购买方区域未识别")
        elif not buyer_hit:
            issues.append("购买方区域未命中白名单")

    seller_name = _find_whitelist_in_text(seller_text, seller_whitelist_map)
    seller_hit = bool(seller_name)
    if not seller_name:
        seller_name = _extract_company_from_region_text(seller_text)
        seller_hit = normalize_company(seller_name) in seller_whitelist_map if seller_name else False
        if not seller_name:
            issues.append("销售方区域未识别")
        elif not seller_hit:
            issues.append("销售方区域未命中白名单")

    return clean_company_name(seller_name), seller_hit, clean_company_name(buyer_name), buyer_hit, issues, layout_dump


def detect_invoice_number(text: str) -> str:
    patterns = [
        r"(?:发票号码|票据号码|Invoice\s*No\.?)[:：]?\s*([A-Za-z0-9\-]{6,24})",
        r"No\.?[:：]?\s*([A-Za-z0-9\-]{6,24})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return normalize_invoice_number(match.group(1))

    nums = re.findall(r"\b\d{8,20}\b", text)
    if nums:
        nums.sort(key=len, reverse=True)
        return normalize_invoice_number(nums[0])

    mixed_ids = re.findall(r"\b[A-Za-z]{1,8}\d{4,20}\b", text)
    if mixed_ids:
        mixed_ids.sort(key=len, reverse=True)
        return normalize_invoice_number(mixed_ids[0])
    return ""


def detect_amount(text: str) -> float:
    normalized_text = text.replace(" ", "")
    precise_patterns = [
        r"价税合计（小写）[^0-9¥￥]{0,20}[¥￥]?\s*([0-9]{1,9}(?:,[0-9]{3})*\.[0-9]{2})",
        r"小写[^0-9¥￥]{0,20}[¥￥]?\s*([0-9]{1,9}(?:,[0-9]{3})*\.[0-9]{2})",
        r"价税合计[^0-9¥￥]{0,20}[¥￥]?\s*([0-9]{1,9}(?:,[0-9]{3})*\.[0-9]{2})",
        r"(?:amount|total)[^0-9¥￥]{0,20}[¥￥]?\s*([0-9]{1,9}(?:,[0-9]{3})*\.[0-9]{2})",
    ]
    for pattern in precise_patterns:
        match = re.search(pattern, normalized_text, re.IGNORECASE)
        if match:
            return float(match.group(1).replace(",", ""))

    nums = re.findall(r"\b([0-9]{1,9}(?:,[0-9]{3})*(?:\.[0-9]{1,2}))\b", text)
    floats = []
    for value in nums:
        try:
            numeric = float(value.replace(",", ""))
        except ValueError:
            continue
        if 0 < numeric < 1_000_000_000:
            floats.append(numeric)
    return max(floats) if floats else 0.0


def detect_invoice_date(text: str) -> str:
    patterns = [
        r"开票日期[:：]?\s*(\d{4}[年\-/.]\d{1,2}[月\-/.]\d{1,2}日?)",
        r"日期[:：]?\s*(\d{4}[年\-/.]\d{1,2}[月\-/.]\d{1,2}日?)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return normalize_date(match.group(1))

    match = re.search(r"(20\d{2}[年\-/.\s]\d{1,2}[月\-/.\s]\d{1,2}日?)", text)
    if match:
        return normalize_date(match.group(1))
    return ""


def _should_try_ocr(
    text: str,
    seller_name: str,
    buyer_name: str,
    invoice_number: str,
    amount: float,
    invoice_date: str,
) -> bool:
    if len(text.strip()) < 30:
        return True
    if not seller_name or not buyer_name or not invoice_number or not invoice_date:
        return True
    if amount <= 0:
        return True
    return False


def parse_invoice(
    pdf_path: Path,
    seller_whitelist_map: dict[str, str],
    buyer_whitelist_map: dict[str, str],
) -> ParsedInvoice:
    base_text = read_pdf_text(pdf_path)
    invoice_number = detect_invoice_number(base_text)
    amount = detect_amount(base_text)
    invoice_date = detect_invoice_date(base_text)
    merged_text = base_text
    ocr_used = False

    seller_name = ""
    buyer_name = ""
    seller_hit = False
    buyer_hit = False
    issues: list[str] = []

    if _should_try_ocr(base_text, seller_name, buyer_name, invoice_number, amount, invoice_date):
        try:
            tokens, ocr_text = read_pdf_ocr_tokens(pdf_path, page_index=0, scale=3.2)
            if ocr_text.strip():
                ocr_used = True
                seller_name, seller_hit, buyer_name, buyer_hit, region_issues, layout_dump = _extract_parties_from_fixed_layout(
                    tokens,
                    seller_whitelist_map,
                    buyer_whitelist_map,
                )
                issues.extend(region_issues)
                merged_text = "\n".join(part for part in [base_text.strip(), ocr_text.strip(), layout_dump.strip()] if part)

                if not invoice_number:
                    invoice_number = detect_invoice_number(merged_text)
                if not invoice_date:
                    invoice_date = detect_invoice_date(merged_text)
                ocr_amount = detect_amount(merged_text)
                if amount <= 0 and ocr_amount > 0:
                    amount = ocr_amount
                elif amount > 100000 and 0 < ocr_amount < amount:
                    amount = ocr_amount
        except Exception:
            issues.append("OCR解析异常")

    if not ocr_used:
        merged_text = base_text

    if not is_valid_company_name(seller_name):
        seller_name = ""
        issues.append("销售方无效")
    elif not seller_hit:
        issues.append("新销售方")

    if not is_valid_company_name(buyer_name):
        buyer_name = ""
        issues.append("购买方无效")
    elif not buyer_hit:
        issues.append("新购买方")

    if seller_name and buyer_name and normalize_company(seller_name) == normalize_company(buyer_name):
        issues.append("销售方与购买方相同")

    if not invoice_number:
        invoice_number = normalize_invoice_number(pdf_path.stem)
        issues.append("发票号码缺失")
    if not invoice_date:
        issues.append("开票日期缺失")
    if amount <= 0:
        issues.append("金额识别异常")

    issues = list(dict.fromkeys(issues))

    return ParsedInvoice(
        seller_name=seller_name,
        buyer_name=buyer_name,
        invoice_number=invoice_number,
        amount=amount,
        invoice_date=invoice_date,
        raw_text=merged_text,
        seller_in_whitelist=seller_hit,
        buyer_in_whitelist=buyer_hit,
        ocr_used=ocr_used,
        issues=issues,
    )
