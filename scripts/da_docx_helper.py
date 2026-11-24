#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Công cụ hỗ trợ kiểm tra khung đề mục, định dạng chuẩn Báo cáo Bài tập lớn (BTL)
& Đồ án tốt nghiệp (ĐATN - KMA), rà soát dấu hiệu văn phong AI Tiếng Việt
(dựa trên dữ liệu thực nghiệm ViDetect & VietAIDetector) và ghi trực tiếp
nội dung / khung giữ chỗ hình ảnh (style Hnhv, Tnhnhv) vào file .docx.
Sử dụng thư viện chuẩn của Python (không bắt buộc cài gói ngoài).
"""

import argparse
import io
import os
import re
import shutil
import stat
import struct
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_URI}
W_VAL = f"{{{W_URI}}}val"
W_ID = f"{{{W_URI}}}styleId"
W_TYPE = f"{{{W_URI}}}type"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MC_IGN = f"{{{MC_NS}}}Ignorable"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Đăng ký đầy đủ các namespace chuẩn của WordprocessingML / DrawingML (Office 2007 - 365)
# để tuyệt đối không sinh tiền tố giả ns0, ns1..ns8 làm Word báo lỗi khôi phục file.
WORD_NAMESPACES = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "mc": MC_NS,
    "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "o": "urn:schemas-microsoft-com:office:office",
    "oel": "http://schemas.microsoft.com/office/2019/extlst",
    "r": R_NS_DOC,
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp15": "http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing",
    "wp": WP_NS,
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": W_URI,
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16cur": "http://schemas.microsoft.com/office/word/2018/wordml/cur",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "a": A_NS,
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
    "pic": PIC_NS,
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
for _prefix, _uri in WORD_NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)

# Thứ tự chuẩn XSD CT_PPrBase (ISO/IEC 29500-1) cho các thẻ con bên trong <w:pPr>
PPR_ELEMENT_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]
_PPR_RANK = {name: idx for idx, name in enumerate(PPR_ELEMENT_ORDER)}


def _sort_ppr_children(ppr: ET.Element):
    """Sắp xếp các thẻ con của <w:pPr> đúng thứ tự XSD chuẩn để Word không báo lỗi cấu trúc XML."""
    children = list(ppr)
    if len(children) <= 1:
        return
    sorted_children = sorted(
        children,
        key=lambda el: _PPR_RANK.get(el.tag.split("}")[-1], len(PPR_ELEMENT_ORDER))
        if el.tag.startswith(f"{{{W_URI}}}")
        else len(PPR_ELEMENT_ORDER) + 1,
    )
    if sorted_children != children:
        for ch in children:
            ppr.remove(ch)
        for ch in sorted_children:
            ppr.append(ch)


def _sort_all_pprs(root: ET.Element):
    for ppr in root.findall(".//w:pPr", NS):
        _sort_ppr_children(ppr)


def parse_ooxml_part(xml_bytes: bytes):
    """
    Phân tích phần tử XML trong gói .docx đồng thời thu nhận toàn bộ cặp (prefix, uri) gốc
    để bảo toàn khai báo xmlns cho mc:Ignorable và tránh sinh tiền tố giả ns0, ns1.
    """
    ns_map = dict(WORD_NAMESPACES)
    try:
        for _event, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
            if prefix:
                ns_map[prefix] = uri
                if uri not in ET._namespace_map:
                    ET.register_namespace(prefix, uri)
    except ET.ParseError:
        pass
    return ET.fromstring(xml_bytes), ns_map


def serialize_ooxml_part(root: ET.Element, ns_map: dict = None, default_ns: str = None) -> bytes:
    """
    Tuần tự hóa cây XML chuẩn Microsoft Word:
    - Cô lập trạng thái ET._namespace_map để default_ns của .rels / [Content_Types].xml không rò rỉ.
    - Bổ sung đầy đủ khai báo xmlns:<prefix> cho mọi tiền tố có mặt trong mc:Ignorable.
    - Xuất tiêu đề chuẩn <?xml version="1.0" encoding="UTF-8" standalone="yes"?>.
    """
    merged_ns = dict(WORD_NAMESPACES)
    if ns_map:
        merged_ns.update(ns_map)

    saved_ns_map = dict(ET._namespace_map)
    try:
        if default_ns:
            ET.register_namespace("", default_ns)
        else:
            for k, v in list(ET._namespace_map.items()):
                if v == "":
                    del ET._namespace_map[k]
            for pfx, uri in WORD_NAMESPACES.items():
                ET._namespace_map[uri] = pfx
        xml_str = ET.tostring(root, encoding="unicode")
    finally:
        ET._namespace_map.clear()
        ET._namespace_map.update(saved_ns_map)

    # Đảm bảo mọi prefix liệt kê trong mc:Ignorable đều có khai báo xmlns:<prefix> trên thẻ gốc
    ign_matches = re.findall(r'\bmc:Ignorable="([^"]*)"', xml_str)
    if ign_matches:
        m_open = re.match(r"^(<[A-Za-z0-9_:\-]+)([^>]*?)(/?\s*>)", xml_str, flags=re.DOTALL)
        if m_open:
            tag_head, attrs_str, tag_tail = m_open.group(1), m_open.group(2), m_open.group(3)
            declared = set(re.findall(r'\bxmlns:([A-Za-z0-9_\-]+)=', attrs_str))
            needed_decls = []
            unknown_pfxs = set()
            for grp in ign_matches:
                for pfx in grp.split():
                    if pfx in declared:
                        continue
                    uri = merged_ns.get(pfx)
                    if uri:
                        needed_decls.append(f' xmlns:{pfx}="{uri}"')
                        declared.add(pfx)
                    else:
                        unknown_pfxs.add(pfx)
            if needed_decls:
                new_open = tag_head + attrs_str + "".join(needed_decls) + tag_tail
                xml_str = new_open + xml_str[m_open.end():]
            if unknown_pfxs:
                def _clean_ign(match_obj):
                    kept = [x for x in match_obj.group(1).split() if x not in unknown_pfxs]
                    return f'mc:Ignorable="{" ".join(kept)}"' if kept else ""
                xml_str = re.sub(r'\bmc:Ignorable="([^"]*)"', _clean_ign, xml_str)

    header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return (header + xml_str).encode("utf-8")


def write_clean_zip_entry(zout: zipfile.ZipFile, filename: str, data: bytes, date_time=None):
    """Ghi một thành phần vào gói .docx bằng ZipInfo sạch (tránh giữ lại flag_bits/extra lỗi từ luồng cũ)."""
    zi = zipfile.ZipInfo(filename=filename, date_time=date_time or time.localtime()[:6])
    zi.compress_type = zipfile.ZIP_DEFLATED
    zi.create_system = 0
    zi.external_attr = 0o600 << 16
    zout.writestr(zi, data)


def ensure_docx_writable_and_replace(tmp_docx: str, docx_path: str):
    """
    Kiểm tra khóa file của Word, gỡ cờ Read-Only và luồng Zone.Identifier (tránh Protected View
    bắt Save As thành file mới), sau đó ghi đè nguyên tử lên file đích.
    """
    folder = os.path.dirname(os.path.abspath(docx_path))
    base = os.path.basename(docx_path)
    lock_candidates = [
        os.path.join(folder, "~$" + base),
        os.path.join(folder, "~$" + base[2:]) if len(base) > 2 else "",
    ]
    for lf in lock_candidates:
        if lf and os.path.exists(lf):
            if os.path.exists(tmp_docx):
                os.remove(tmp_docx)
            raise RuntimeError(
                f"File '{docx_path}' đang được mở trong Microsoft Word (tồn tại file khóa '{os.path.basename(lf)}'). "
                "Vui lòng đóng cửa sổ Word trước khi chạy cập nhật để không bị xung đột lưu file."
            )

    if os.path.exists(docx_path):
        try:
            os.chmod(docx_path, stat.S_IWRITE | stat.S_IREAD)
        except OSError:
            pass
        try:
            with open(docx_path, "r+b"):
                pass
        except PermissionError as exc:
            if os.path.exists(tmp_docx):
                os.remove(tmp_docx)
            raise RuntimeError(
                f"Không thể ghi vào '{docx_path}' vì file đang bị khóa bởi tiến trình khác (ví dụ Microsoft Word). "
                "Hãy đóng file trong Word rồi chạy lại."
            ) from exc

    os.replace(tmp_docx, docx_path)
    try:
        os.chmod(docx_path, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass
    # Xóa Alternate Data Stream Zone.Identifier trên Windows (nếu file tải từ mạng về khiến Word bật Protected View)
    if os.name == "nt":
        ads_path = docx_path + ":Zone.Identifier"
        if os.path.exists(ads_path):
            try:
                os.remove(ads_path)
            except OSError:
                pass


def get_image_pixel_dimensions(img_bytes: bytes):
    """Đọc kích thước điểm ảnh (width, height) của ảnh PNG hoặc JPEG bằng thư viện chuẩn Python."""
    if len(img_bytes) >= 24 and img_bytes[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", img_bytes[16:24])
        if w > 0 and h > 0:
            return w, h
    if len(img_bytes) > 4 and img_bytes[:2] == b"\xff\xd8":
        idx = 2
        n = len(img_bytes)
        while idx + 8 < n:
            if img_bytes[idx] != 0xFF:
                idx += 1
                continue
            marker = img_bytes[idx + 1]
            while marker == 0xFF and idx + 2 < n:
                idx += 1
                marker = img_bytes[idx + 1]
            if marker in (0xD8, 0xD9) or (0xD0 <= marker <= 0xD7):
                idx += 2
                continue
            if idx + 4 > n:
                break
            seg_len = struct.unpack(">H", img_bytes[idx + 2 : idx + 4])[0]
            if seg_len < 2 or idx + 2 + seg_len > n:
                break
            if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                if idx + 9 <= n:
                    h, w = struct.unpack(">HH", img_bytes[idx + 5 : idx + 9])
                    if w > 0 and h > 0:
                        return w, h
                break
            idx += 2 + seg_len
    return None


def validate_docx_ooxml(docx_path: str) -> list:
    """Kiểm tra tính toàn vẹn cấu trúc OOXML để đảm bảo Microsoft Word mở và Ctrl+S trực tiếp không lỗi."""
    problems = []
    if not os.path.exists(docx_path):
        return [f"Không tìm thấy file: {docx_path}"]
    if not os.access(docx_path, os.W_OK):
        problems.append("[Quyền ghi File] File đang ở trạng thái Read-Only khiến Word bắt Save As thành tài liệu mới")
    if os.name == "nt" and os.path.exists(docx_path + ":Zone.Identifier"):
        problems.append("[Protected View] File còn gắn luồng Zone.Identifier khiến Word mở ở chế độ Protected View")

    with zipfile.ZipFile(docx_path, "r") as z:
        for name in z.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            raw = z.read(name)
            txt = raw.decode("utf-8", errors="replace")
            try:
                root = ET.fromstring(raw)
            except ET.ParseError as exc:
                problems.append(f"[Lỗi cú pháp XML] Thành phần '{name}' không hợp lệ: {exc}")
                continue

            ns_fake = sorted(set(re.findall(r"</?(ns\d+):|\bxmlns:(ns\d+)=", txt)))
            if ns_fake:
                problems.append(f"[Lỗi Namespace OOXML] '{name}' chứa tiền tố giả {ns_fake} dễ gây hộp thoại khôi phục Word")

            mc_ign = re.findall(r'\bmc:Ignorable="([^"]+)"', txt)
            if mc_ign:
                m_open = re.match(r"^<[A-Za-z0-9_:\-]+([^>]*?)>", re.sub(r"^<\?xml[^?]*\?>\s*", "", txt), flags=re.DOTALL)
                decl = set(re.findall(r'\bxmlns:([A-Za-z0-9_\-]+)=', m_open.group(1) if m_open else txt[:3000]))
                missing = [p for g in mc_ign for p in g.split() if p not in decl]
                if missing:
                    problems.append(
                        f"[Lỗi mc:Ignorable] '{name}' khai báo mc:Ignorable thiếu xmlns cho {missing} -> Word sẽ báo hỏng file và bắt Save As"
                    )

            if name in ("word/document.xml", "word/styles.xml"):
                bad_ppr = 0
                for ppr in root.findall(".//w:pPr", NS):
                    ranks = [
                        _PPR_RANK[c.tag.split("}")[-1]]
                        for c in ppr
                        if c.tag.startswith(f"{{{W_URI}}}") and c.tag.split("}")[-1] in _PPR_RANK
                    ]
                    if ranks != sorted(ranks):
                        bad_ppr += 1
                if bad_ppr:
                    problems.append(f"[Lỗi thứ tự XSD] '{name}' có {bad_ppr} thẻ <w:pPr> sai thứ tự phần tử con chuẩn OOXML")

            if name == "word/document.xml":
                doc_prs = [el.attrib.get("id") for el in root.findall(f".//{{{WP_NS}}}docPr") if el.attrib.get("id")]
                if len(doc_prs) != len(set(doc_prs)):
                    problems.append(f"[Trùng ID ảnh DrawingML] 'word/document.xml' có wp:docPr trùng id ({len(doc_prs)} phần tử)")
    return problems


def make_w_el(local_name: str, attrib: dict = None) -> ET.Element:
    """Tạo phần tử WordprocessingML (prefix w:) kèm các thuộc tính tương ứng."""
    el = ET.Element(f"{{{W_URI}}}{local_name}")
    if attrib:
        for k, v in attrib.items():
            if k.startswith("{"):
                el.set(k, v)
            else:
                el.set(f"{{{W_URI}}}{k}", v)
    return el

# ==============================================================================
# 1. BỘ QUY TẮC TIỀN XỬ LÝ TIẾNG VIỆT (Trích trực tiếp từ VietAIDetector)
#    Nguồn: VietAIDetector-main/preprocessing/text_utils.py & normalizer.py
# ==============================================================================
DEHYPHEN_RE = re.compile(r"([a-zA-ZÀ-ỹà-ỹ])-\s*\n\s*([a-zA-ZÀ-ỹà-ỹ])", re.UNICODE)
LINEBREAK_RE = re.compile(r"(?<![.!?…])\n(?!\n)")
WHITESPACE_RE = re.compile(r"[ \t\r\u00a0\ufeff]+")
NON_SEMANTIC_RE = re.compile(
    r"[^\w\sÀ-ỹà-ỹ.,!?;:()\-\'\"“”‘’]",
    re.UNICODE,
)


def count_words_vi(text: str) -> int:
    return len(text.split())


def special_char_ratio_vi(text: str) -> float:
    if not text:
        return 0.0
    return len(NON_SEMANTIC_RE.findall(text)) / len(text)


def normalize_vietnamese_paragraphs(raw_text: str):
    """Chuẩn hóa đoạn văn theo đúng pipeline TextNormalizer của VietAIDetector."""
    dehyphened = DEHYPHEN_RE.sub(r"\1\2", raw_text)
    valid_paragraphs = []
    for p in dehyphened.split("\n"):
        p_norm = WHITESPACE_RE.sub(" ", p).strip()
        if not p_norm:
            continue
        if count_words_vi(p_norm) < 5:
            continue
        if special_char_ratio_vi(p_norm) > 0.50:
            continue
        p_clean = LINEBREAK_RE.sub(" ", p_norm)
        p_clean = WHITESPACE_RE.sub(" ", p_clean).strip()
        valid_paragraphs.append(p_clean)
    return valid_paragraphs


# ==============================================================================
# 2. MẪU DẤU HIỆU VĂN PHONG AI TIẾNG VIỆT
#    Trích xuất thực nghiệm từ 7.222 cặp văn bản source.jsonl và 120 văn bản
#    evaluation.json (Claude-Sonnet-4.6, Gemini-3.7-Flash, GPT-5.6-Luna)
#    trong VietAIDetector-main.zip + quy chuẩn ĐA KMA.
# ==============================================================================
AI_TELL_PATTERNS_VI = [
    # Nhóm 1: Dư lượng dẫn nhập / hội thoại AI, giới hạn kiến thức & dấu vết công cụ
    (
        r"(?i)(?:^(?:dưới\s+đây\s+là\s+(?:nội\s+dung|bản|bài|phần\s+trình\s+bày)|đây\s+là\s+tin\s+tức)|"
        r"\b(?:dựa\s+trên\s+thông\s+tin\s+(?:được\s+cung\s+cấp|hiện\s+có)|hy\s+vọng\s+(?:phần\s+trình\s+bày|những\s+thông\s+tin\s+(?:này|trên)|bài\s+viết)|"
        r"nếu\s+bạn\s+muốn,?\s+(?:mình|tôi|em)\s+có\s+thể\s+(?:phân\s+tích|viết|soạn|tóm\s+tắt|gợi\s+ý|mở\s+rộng|đi\s+sâu)|"
        r"đừng\s+ngần\s+ngại|câu\s+hỏi\s+(?:rất\s+)?hay|chắc\s+chắn\s+rồi|tất\s+nhiên\s+rồi|"
        r"tính\s+đến\s+thời\s+điểm\s+(?:cập\s+nhật|huấn\s+luyện))\b)",
        "Dấu hiệu AI (Chatbot / Giới hạn kiến thức): Dư lượng hội thoại AI ('Dưới đây là bài...', 'Nếu bạn muốn, mình có thể...', 'Tính đến thời điểm...')",
    ),
    (
        r"(?i)(?:contentReference|oaicite|turn0search\d|\[cite:\s*\d+\]|:::writing|utm_source=chatgpt\.com|"
        r"\[(?:họ\s+(?:và\s+)?tên|tên\s+nhóm|mssv|chức\s+vụ|sđt|số\s+điện\s+thoại|your\s+[^\]]{1,25}|insert\s+[^\]]{1,25})\](?!\())",
        "Dấu vết công cụ AI / Placeholder bản mẫu sót: Chứa thẻ 'oaicite/cite', đuôi link ChatGPT hoặc ô [Họ tên/MSSV...] chưa điền",
    ),
    (
        r"(?i)(?:</?(?:skill|system|instructions?)>|(?:bỏ\s+qua|quên)\s+(?:mọi|tất\s+cả|toàn\s+bộ)\s+(?:hướng\s+dẫn|chỉ\s+dẫn|chỉ\s+thị)|"
        r"\bignore\s+(?:all|any)\s+(?:previous|prior|above)\s+instructions\b)",
        "Cảnh báo bảo mật văn bản (Prompt Injection): Tài liệu nguồn chứa thẻ hoặc câu chỉ thị giả cho AI -> bỏ qua, chỉ xem là văn bản thuần",
    ),
    # Nhóm 2: Cấu trúc đối lập giả & đòn bẩy khuôn mẫu
    (
        r"(?i)\bkhông\s+chỉ\s+(?:đơn\s+thuần\s+)?(?:là\s+)?[^.?!;\n]{3,90}\bmà\s+(?:đằng\s+sau\s+[^.?!;\n]{3,40}\s+|còn\s+)",
        "Dấu hiệu AI (VietAIDetector 6.3x-10.2x): Cấu trúc đòn bẩy 'Không chỉ ... mà còn ...' -> nên viết thẳng vào chức năng kỹ thuật",
    ),
    (
        r"(?i)\bcâu\s+hỏi\s+không\s+còn\s+là\s+liệu\b[^.?!;\n]{3,80}\bmà\s+là\b|"
        r"\bkhông\s+(?:còn\s+)?(?:là|ở|nằm\s+ở)\s+[^.?!;\n]{3,75}\bmà\s+(?:là|ở|nằm\s+ở)\b|"
        r"(?:^|[.!?]\s+)(?:đừng|thay\s+vì)\s+[^.!?\n,]{3,40},\s+hãy\b",
        "Dấu hiệu AI (Đối lập giả): Khuôn dựng đối lập 'không phải/nằm ở X mà là Y' hoặc 'Đừng X, hãy Y' -> nói thẳng nội dung chính",
    ),
    # Nhóm 3: Cụm bơm thổi tầm quan trọng thực nghiệm từ source.jsonl & evaluation.json
    (
        r"(?i)\bnhấn\s+mạnh\s+tầm\s+quan\s+trọng\b",
        "Dấu hiệu AI (VietAIDetector 135x): Cụm sáo rỗng 'nhấn mạnh tầm quan trọng'",
    ),
    (
        r"(?i)\blà\s+một\s+bước\s+(tiến|ngoặt)\s+quan\s+trọng\b",
        "Dấu hiệu AI (VietAIDetector 93x): Cụm bơm thổi 'là một bước tiến/bước ngoặt quan trọng'",
    ),
    (
        r"(?i)\bđã\s+làm\s+dấy\s+lên\b",
        "Dấu hiệu AI (VietAIDetector 94x): Cụm dịch máy 'đã làm dấy lên'",
    ),
    (
        r"(?i)\bđóng\s+vai\s+trò\s+(như\s+một\s+)?(then\s+chốt|sống\s+còn|xương\s+sống|vô\s+cùng\s+quan\s+trọng)\b",
        "Dấu hiệu AI (VietAIDetector 12x-20.5x): Cụm bơm thổi 'đóng vai trò then chốt/sống còn/xương sống'",
    ),
    (
        r"(?i)\b(điểm\s+đáng\s+chú\s+ý\s+là|đáng\s+chú\s+ý\s+là)\b",
        "Dấu hiệu AI (VietAIDetector evaluation.json 21x): Cụm mở câu khuôn mẫu 'Điểm đáng chú ý là...'",
    ),
    (
        r"(?i)(?:^|[.!?]\s+)(?:nói\s+(?:cho\s+)?(?:gọn|thật|thẳng|rõ|công\s+bằng)|thành\s+thật\s+mà\s+nói|công\s+bằng\s+mà\s+nói|hiểu\s+(?:một\s+cách\s+)?(?:đơn\s+giản|nôm\s+na)(?:\s+nhất)?)\s*(?:là|thì)?[:,]",
        "Dấu hiệu AI (Nhãn báo giọng): Cụm rào trước kiểu 'Nói gọn:/Nói thật:/Hiểu đơn giản là:' -> bỏ nhãn, đi thẳng vào ý",
    ),
    # Nhóm 4: Mẫu câu bình luận vô thưởng vô phạt đặc trưng của GPT/Claude/Gemini trong evaluation.json
    (
        r"(?i)\bkhông\s+thể\s+được\s+nhìn\s+nhận\s+chỉ\s+từ\s+một\s+chiều\s+cạnh\b",
        "Dấu hiệu AI (GPT-5.6 benchmark 105x): Câu rào đón 'không thể được nhìn nhận chỉ từ một chiều cạnh'",
    ),
    (
        r"(?i)\b(một\s+lựa\s+chọn\s+hợp\s+lý\s+thường\s+phải\s+cân\s+bằng\s+giữa|sự\s+cân\s+bằng\s+này\s+quyết\s+định\s+đáng\s+kể\s+đến\s+khả\s+năng\s+duy\s+trì)\b",
        "Dấu hiệu AI (GPT-5.6 benchmark 106x): Câu bình luận khuôn mẫu 'phải cân bằng giữa... quyết định đáng kể đến khả năng duy trì...'",
    ),
    (
        r"(?i)\bviệc\s+đánh\s+giá\s+cần\s+dựa\s+trên\s+nhiều\s+dấu\s+hiệu\s+thay\s+vì\s+một\s+chỉ\s+báo\s+duy\s+nhất\b",
        "Dấu hiệu AI (GPT-5.6 benchmark 103x): Câu bình luận chung chung 'việc đánh giá cần dựa trên nhiều dấu hiệu thay vì một chỉ báo duy nhất'",
    ),
    (
        r"(?i)\b(khi\s+đặt\s+vấn\s+đề\s+trong\s+một\s+bối\s+cảnh\s+rộng\s+hơn|trong\s+bối\s+cảnh\s+thế\s+giới\s+đang\s+phải\s+đối\s+mặt\s+với\s+những\s+thách\s+thức)\b",
        "Dấu hiệu AI (Claude/GPT benchmark 104x): Mở đoạn sáo rỗng 'Khi đặt vấn đề trong một bối cảnh rộng hơn / Trong bối cảnh thế giới đang phải đối mặt...'",
    ),
    (
        r"(?i)\b(luôn\s+được\s+ví\s+như\s+một\s+thước\s+đo\s+nhạy\s+bén|nổi\s+lên\s+như\s+một\s+hướng\s+đi\s+tất\s+yếu|lựa\s+chọn\s+chiến\s+lược\s+không\s+thể\s+trì\s+hoãn)\b",
        "Dấu hiệu AI (Claude/Gemini benchmark): Cụm văn hoa 'được ví như thước đo nhạy bén / nổi lên như hướng đi tất yếu'",
    ),
    # Nhóm 5: Từ vựng sáo rỗng & Văn dịch tiếng Anh (VietAIDetector + viet-nhu-nguoi Nhóm G/K)
    (
        r"(?i)\bbức\s+tranh\s+toàn\s+cảnh\b",
        "Dấu hiệu AI: Cụm dịch máy 'bức tranh toàn cảnh' (landscape) -> thay bằng 'tổng quan'",
    ),
    (
        r"(?i)\b(một\s+cách\s+toàn\s+diện\s+và\s+sâu\s+sắc|phong\s+phú\s+và\s+sâu\s+sắc)\b",
        "Dấu hiệu AI (VietAIDetector 7.1x): Cụm tán dương 'toàn diện và sâu sắc'",
    ),
    (
        r"(?i)\b(tối\s+ưu\s+hóa\s+vượt\s+trội|khai\s+phá\s+tiềm\s+năng|hành\s+trình\s+khám\s+phá|mở\s+khóa\s+tiềm\s+năng|nâng\s+tầm|giá\s+trị\s+cốt\s+lõi|tạo\s+ra\s+giá\s+trị\s+thực|góp\s+(?:một\s+)?phần\s+không\s+nhỏ)\b",
        "Dấu hiệu AI (Quảng cáo / Văn dịch): Cụm sáo rỗng 'tối ưu hóa vượt trội / mở khóa tiềm năng / nâng tầm / giá trị cốt lõi'",
    ),
    (
        r"(?i)\b(điều\s+hướng\s+(?:sự|những)\s+phức\s+tạp|tác\s+động\s+có\s+ý\s+nghĩa|(?:ở|tại)\s+cấp\s+độ\s+cốt\s+lõi|trải\s+nghiệm\s+liền\s+mạch|"
        r"khi\s+nói\s+đến|ở\s+cuối\s+ngày|đạt\s+được\s+những\s+tiến\s+bộ\s+đáng\s+kể|tiến\s+về\s+phía\s+trước|"
        r"không\s+phải\s+là\s+không\s+có\s+(?:những\s+)?(?:hạn\s+chế|nhược\s+điểm|thách\s+thức|rủi\s+ro)|một\s+loạt\s+các|"
        r"là\s+(?:rất|vô\s+cùng)\s+(?:quan\s+trọng|cần\s+thiết|hữu\s+ích)\s+(?:để|rằng)|"
        r"\b(?:nơi|người|điều|thứ|lý\s+do)\s+mà\b(?!\s+(?:không|chưa|vẫn|chẳng|cũng|còn|lại)\b)|\bcũng\s+như\s+là\b|\bvà/hoặc\b)",
        "Dấu hiệu văn dịch (viet-nhu-nguoi Nhóm K): Cụm dịch sát tiếng Anh ('trải nghiệm liền mạch / khi nói đến / một loạt các / điều mà / và/hoặc...') -> viết thuần Việt tự nhiên",
    ),
    (
        r"(?i)(?:\bđược\s+(?:thực\s+hiện|ghi\s+nhận|chứng\s+kiến)\s+bởi\b|(?<!được\s)(?<!bị\s)\b(?:dẫn\s+dắt|dẫn\s+đầu|tài\s+trợ|xây\s+dựng|thiết\s+kế|tạo\s+ra)\s+bởi\b)",
        "Dấu hiệu AI (VietAIDetector 3.75x): Câu bị động dịch thô '... bởi' -> chuyển sang câu chủ động nêu rõ chủ thể",
    ),
    (
        r"(?i)\bđó\s+chính\s+là\s+(chìa\s+khóa|mấu\s+chốt)\b",
        "Dấu hiệu AI: Câu chốt kịch tính 'Đó chính là chìa khóa/mấu chốt'",
    ),
    # Nhóm 6: Lỗi dẫn nguồn lộ liễu từ tài liệu đính kèm
    (
        r"(?i)\b(theo\s+tài\s+liệu\s+(đính\s+kèm|cung\s+cấp|tham\s+khảo)|tài\s+liệu\s+(này|trên|đính\s+kèm)\s+cho\s+biết|bài\s+báo\s+(này|trên)\s+cho\s+biết|theo\s+slide\s+bài\s+giảng|như\s+trong\s+file\s+đính\s+kèm)\b",
        "Lỗi dẫn nguồn rập khuôn: Không viết 'theo tài liệu đính kèm / bài báo này cho biết...'; hãy diễn đạt thẳng nội dung thành lời văn tự nhiên của báo cáo",
    ),
    # Nhóm 7: Danh hóa chồng chất, động từ rỗng, thừa nghĩa lặp & rào đón (human-main + viet-nhu-nguoi)
    (
        r"(?i)\b(?:việc\s+)?tiến\s+hành\s+thực\s+hiện\b|\b(?:cung\s+cấp|đưa\s+ra)\s+sự\s+(?:hỗ\s+trợ|giúp\s+đỡ)\b|\bhiện\s+thực\s+hóa\b",
        "Dấu hiệu AI (Động từ rỗng / Danh hóa): Rút gọn 'tiến hành thực hiện / cung cấp sự hỗ trợ / hiện thực hóa' thành động từ chính",
    ),
    (
        r"(?i)\bnhằm\s+mục\s+đích\s+để\b|\b(?:hiện\s+nay\s+đang|đang\s+trong\s+quá\s+trình|hoàn\s+toàn\s+đầy\s+đủ|các\s+những|mọi\s+tất\s+cả|vẫn\s+còn\s+đang|quay\s+trở\s+lại|tái\s+[a-zà-ỹ]+\s+lại)\b",
        "Lỗi thừa nghĩa lặp (viet-nhu-nguoi Mục 35): Rút gọn cụm thừa ('nhằm mục đích để / hiện nay đang / quay trở lại / vẫn còn đang...')",
    ),
    (
        r"(?i)\bcủa\s+[^,.;:?!()\n]{1,25}\bcủa\s+[^,.;:?!()\n]{1,25}\bcủa\b",
        "Dấu hiệu AI (human-main Điều 1b): Chuỗi 'của... của... của...' lồng nhau kiểu dịch tiếng Anh -> viết lại câu chủ động",
    ),
    (
        r"(?i)\b(không\s+thể\s+phủ\s+nhận\s+rằng|như\s+chúng\s+ta\s+đã\s+biết|có\s+thể\s+khẳng\s+định\s+rằng|không\s+thể\s+không\s+nhắc\s+đến|cần\s+lưu\s+ý\s+rằng\s+là)\b",
        "Dấu hiệu AI (human-main Điều 5): Cụm rào đón thừa ('Không thể phủ nhận rằng / Như chúng ta đã biết...') -> đi thẳng vào luận điểm",
    ),
    # Nhóm 8: Bổ sung từ viet-style-main (V03, V04: Dấu phẩy trước 'và', ký hiệu '&' trong văn xuôi tiếng Việt)
    (
        r",\s+và\s+[a-zà-ỹ]",
        "Lỗi dấu câu Tiếng Việt (viet-style V04): Không đặt dấu phẩy trước chữ 'và' (', và ' -> ' và ')",
    ),
    (
        r"[A-ZÀ-Ỹa-zà-ỹ]*[À-ỹà-ỹ][A-ZÀ-Ỹa-zà-ỹ]*\s+&\s+[A-ZÀ-Ỹa-zà-ỹ]*[À-ỹà-ỹ]",
        "Lỗi ký tự Tiếng Việt (viet-style V03): Trong câu văn tiếng Việt dùng chữ 'và' thay cho ký hiệu '&'",
    ),
    # Nhóm 9: Từ ngữ hoa mỹ không tự nhiên cho văn kỹ thuật & chú thích tiếng Anh thừa cho từ phổ thông
    (
        r"(?i)\b(giải\s+phẫu\s+cơ\s+chế|bóc\s+tách|nhược\s+điểm\s+chí\s+mạng|bức\s+tranh\s+đầy\s+đủ|sự\s+cộng\s+hưởng|đè\s+nặng\s+lên)\b",
        "Dấu hiệu AI (văn phong hoa mỹ/kịch tính): Tránh dùng 'giải phẫu cơ chế / bóc tách / nhược điểm chí mạng / cộng hưởng / đè nặng' trong báo cáo kỹ thuật",
    ),
    (
        r"(?i)\s+\((?:firewall|firewalls|server|servers|packets|header|bot|bots|routers|switches|load\s+balancers|ping\s+request|form\s+submission)\)",
        "Dấu hiệu AI (dịch thuật máy): Chú thích tiếng Anh thừa trong ngoặc cho thuật ngữ quá quen thuộc -> chỉ giữ tiếng Việt hoặc viết thẳng thuật ngữ",
    ),
    (
        r"(?i)\bgiúp\s+(?:sinh\s+viên|người\s+đọc|chúng\s+ta)\s+(?:nắm|hiểu|hình\s+dung|biết)\b",
        "Dấu hiệu AI (giọng giảng giải khuôn mẫu): Tránh viết 'giúp sinh viên/người đọc nắm/hiểu...' -> nêu trực tiếp mục tiêu khảo sát hoặc tác dụng kỹ thuật của nhóm",
    ),
]

# Các câu cảm tính chỉ cấm nghiêm ngặt ở chế độ Đồ án tốt nghiệp (do-an)
DO_AN_ONLY_PATTERNS = [
    (
        r"(?i)\b(em|chúng\s+em|mình)\s+đã\s+(hiểu|nắm|biết|học\s+hỏi)\s+được\b",
        "Lỗi chuẩn ĐATN KMA (Kết luận): Cấm dùng câu cảm tính 'em/chúng em đã hiểu/nắm/biết được...'",
    ),
    (
        r"(?i)\b(kiến\s+thức\s+thu\s+được\s+rất\s+bổ\s+ích|mang\s+lại\s+rất\s+nhiều\s+kiến\s+thức\s+và\s+kỹ\s+năng\s+làm\s+việc\s+cho\s+chúng\s+em)\b",
        "Lỗi chuẩn ĐATN KMA (Kết luận): Trong ĐATN không viết cảm nghĩ cá nhân; chỉ chấp nhận ở Báo cáo Bài tập lớn (--mode btl)",
    ),
]

PLACEHOLDER_MARKER_RE = re.compile(
    r"(?i)(\[KHUNG\s+GIỮ\s+CHỖ\s+HÌNH|\[KHUNG\s+HÌNH|\[ĐỂ\s+TRỐNG\s+HÌNH\s+ẢNH|\[CHÈN\s+HÌNH)"
)
ASCII_DIAGRAM_RE = re.compile(r"(\+[-\+]{4,}\+|[┌┐└┘├┤┬┴┼─│]{4,}|\[[^\]]+\]\s*--+>\s*\[[^\]]+\])")


# ==============================================================================
# 0. HẰNG SỐ FORMAT CHUẨN ĐỒ ÁN TỐT NGHIỆP (PHỤ LỤC 8 - KMA)
#    Áp dụng thống nhất cho CẢ chế độ Bài tập lớn (btl) và Đồ án tốt nghiệp (do-an)
# ==============================================================================
KMA_PAGE_W = "11906"       # A4: 21.0 cm
KMA_PAGE_H = "16838"       # A4: 29.7 cm
KMA_MARGINS = {
    "top": "1134",         # 2.0 cm
    "bottom": "1134",      # 2.0 cm
    "left": "1701",        # 3.0 cm
    "right": "851",        # 1.5 cm
    "header": "709",       # 1.25 cm
    "footer": "709",       # 1.25 cm
    "gutter": "0",
}


def get_template_path():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.dirname(script_dir), "templates", "Mau_trinh_bay_DA.docx")


def split_sentences_vi(text: str):
    """Tách câu tiếng Việt bảo toàn số thập phân (3.1, 6.1%) và viết tắt."""
    cleaned = re.sub(r"\b(\d+)\.(\d+)", r"\1<DOT>\2", text)
    parts = [s.strip() for s in re.split(r"(?<=[.!?…])\s+", cleaned) if s.strip()]
    return [p.replace("<DOT>", ".") for p in parts if count_words_vi(p) >= 3]


def extract_docx_paragraphs(docx_path):
    """Trích xuất danh sách (index, style_id, text, has_drawing) và căn lề từ file .docx."""
    paragraphs = []
    margins = []
    with zipfile.ZipFile(docx_path, "r") as z:
        xml_bytes = z.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        body = root.find("w:body", NS)
        if body is None:
            return paragraphs, margins
        for idx, p in enumerate(body.findall(".//w:p", NS), 1):
            ppr = p.find("w:pPr", NS)
            style_id = "Normal"
            if ppr is not None:
                ps = ppr.find("w:pStyle", NS)
                if ps is not None:
                    style_id = ps.attrib.get(W_VAL, "Normal")
            texts = [t.text for t in p.findall(".//w:t", NS) if t.text]
            full_text = "".join(texts).strip()
            has_img = (
                p.find(".//w:drawing", NS) is not None
                or p.find(".//w:pict", NS) is not None
                or p.find(".//w:object", NS) is not None
            )
            if full_text or has_img:
                paragraphs.append((idx, style_id, full_text, has_img))
        for pgmar in body.findall(".//w:pgMar", NS):
            margins.append({k.split("}")[-1]: v for k, v in pgmar.attrib.items()})
    return paragraphs, margins


TAKEAWAY_END_SENT_RE = re.compile(
    r"(?i)(\bgiúp\s+(?:người\s+)?(?:quản\s+trị|kỹ\s+sư|đội\s+ngũ|người\s+vận\s+hành|phân\s+biệt)\b|"
    r"^(?:vì\s+vậy|vì\s+thế|bởi\s+vậy|bởi\s+thế|do\s+đó)\b|"
    r"\b(?:đặt\s+nền\s+móng|tạo\s+cơ\s+sở\s+trực\s+tiếp|là\s+yêu\s+cầu\s+thiết\s+thực)\b)"
)
STACCATO_RUN_RE = re.compile(r"(?:(?:^|(?<=[.!?])\s+)(?:[^\s.!?]+\s+){0,3}[^\s.!?]+\.(?=\s|$)){3,}")
COLON_MID_SENT_RE = re.compile(r"(?<!\d)(?<!https)(?<!http):(?!\s*[\"“«‘'–—\-/\d])(?=\s|\S)")

FILLER_WORDS_VI = [
    r"\bcủa\s+(?:bạn|tôi|chúng\s+tôi|chúng\s+ta|họ|anh\s+ấy|cô\s+ấy|nó)\b",
    r"\bmột\b", r"\b(?:các|những)\b", r"\bđược\b", r"\bcó\s+thể\b",
    r"\bsẽ\b", r"\bđã\b", r"\brằng\b", r"\b(?:cho\s+phép|giúp)\b",
    r"\b(?:điều\s+này|nó)\b", r"\b(?:việc|sự)\b",
]
PASSIVE_CALQUE_RE = re.compile(
    r"(?i)\b(?:(?:sẽ|đã|đang|cần|phải|nên|có\s+thể)\s+)?được\s+(?:giữ\s+nguyên|cập\s+nhật|cài\s+đặt|thiết\s+kế|"
    r"xây\s+dựng|phát\s+triển|tạo\s+ra|sử\s+dụng|thực\s+hiện|triển\s+khai|cung\s+cấp|áp\s+dụng|đánh\s+giá|"
    r"xem\s+là|coi\s+là|coi\s+như|lưu\s+trữ|tối\s+ưu|tích\s+hợp|quản\s+lý|xử\s+lý|xác\s+định|lập\s+trình|"
    r"đề\s+cập|thiết\s+lập|cấu\s+hình|ghi\s+nhận|trình\s+bày|mô\s+tả|liệt\s+kê|đặt\s+tên|biết\s+đến)\b(?!\s+bởi)"
)
THING_SUBJECT_RE = re.compile(
    r"(?i)\b(?:bảng|hình|biểu\s+đồ|sơ\s+đồ|kết\s+quả|số\s+liệu|dữ\s+liệu|báo\s+cáo|tài\s+liệu|công\s+cụ|"
    r"tính\s+năng|hệ\s+thống|giải\s+pháp|nền\s+tảng|ứng\s+dụng)(?:\s+\w+){0,3}\s+"
    r"(?:cho\s+phép\s+bạn|giúp\s+bạn|mang\s+lại|mang\s+đến|đảm\s+bảo\s+rằng|cho\s+thấy\s+rằng|chỉ\s+ra\s+rằng)\b"
)


def compute_videtect_stats(paragraph_texts):
    """
    Tính chỉ số cấu trúc đoạn và nhịp độ dài câu trên các đoạn văn xuôi (>= 15 từ)
    kết hợp đối chiếu thực nghiệm trên báo cáo mẫu sinh viên KMA và `viet-nhu-nguoi`:
    - Tỷ lệ đoạn 1-3 câu tự nhiên xen kẽ đoạn phân tích 3-6 câu (cảnh báo nếu >=50% đoạn đóng cứng ở khuôn 4-5 câu).
    - Kiểm soát câu chốt đoạn rập khuôn ('giúp quản trị viên...', 'Vì thế/Vì vậy...').
    - Độ dài câu tự nhiên (~18-28 từ/câu), hệ số biến thiên CV (`std/mean`), chuỗi câu cụt (`staccato`),
      chuỗi câu dài đều nhau và tỷ lệ dấu hai chấm `:` giữa câu.
    """
    prose_paras = []
    sent_counts = []
    word_counts = []
    sent_word_lens = []
    takeaway_end_count = 0
    multi_sent_paras = 0
    staccato_count = 0
    uniform_sent_run_count = 0
    colon_sent_count = 0
    colon_den_sents = 0

    for p in paragraph_texts:
        s = p.strip()
        is_non_prose = (
            PLACEHOLDER_MARKER_RE.search(s)
            or s.startswith(("$", "#", "-", "+", "•", "❖", "Bước ", "Hình ", "Bảng ", "["))
            or re.match(r"^(\(\d+\)|\d+[\.\)]|[a-z]\))\s+", s)
        )
        if not is_non_prose and count_words_vi(s) >= 6 and STACCATO_RUN_RE.search(s):
            staccato_count += 1
        if count_words_vi(s) < 15 or is_non_prose:
            continue
        sents = split_sentences_vi(s)
        if not sents:
            continue
        prose_paras.append(s)
        sent_counts.append(len(sents))
        word_counts.append(count_words_vi(s))

        p_slens = [count_words_vi(st) for st in sents]
        sent_word_lens.extend(p_slens)

        # Kiểm tra 3 câu liên tiếp dài đều nhau (>= 20 từ, chênh nhau <= 4 từ)
        run_eq = 1
        for a, b in zip(p_slens, p_slens[1:]):
            if a >= 20 and b >= 20 and abs(a - b) <= 4:
                run_eq += 1
                if run_eq == 3:
                    uniform_sent_run_count += 1
                    break
            else:
                run_eq = 1

        # Đếm dấu hai chấm giữa câu (trừ câu cuối đoạn dẫn vào danh sách/code/hình)
        lead_in = s.rstrip(" *_").endswith(":")
        body_sents = sents[:-1] if lead_in else sents
        colon_den_sents += len(body_sents)
        colon_sent_count += sum(1 for x in body_sents if COLON_MID_SENT_RE.search(x))

        if len(sents) >= 3:
            multi_sent_paras += 1
            if TAKEAWAY_END_SENT_RE.search(sents[-1]):
                takeaway_end_count += 1

    if not prose_paras or not sent_word_lens:
        return None
    avg_sents_per_para = sum(sent_counts) / len(prose_paras)
    avg_words_per_para = sum(word_counts) / len(prose_paras)
    short_1_3_ratio = sum(1 for c in sent_counts if 1 <= c <= 3) / len(sent_counts)
    uniform_4_5_ratio = sum(1 for c in sent_counts if 4 <= c <= 5) / len(sent_counts)
    takeaway_ratio = (takeaway_end_count / max(multi_sent_paras, 1)) * 100.0
    sorted_slens = sorted(sent_word_lens)
    avg_words_per_sent = sum(sent_word_lens) / len(sent_word_lens)
    median_words_per_sent = sorted_slens[len(sorted_slens) // 2]
    over_45_count = sum(1 for x in sent_word_lens if x > 45)
    over_65_count = sum(1 for x in sent_word_lens if x > 65)
    var_sent = sum((x - avg_words_per_sent) ** 2 for x in sent_word_lens) / len(sent_word_lens)
    std_words_sent = var_sent ** 0.5
    cv_sent_len = round(std_words_sent / avg_words_per_sent, 2) if avg_words_per_sent > 0 else 0.0
    colon_sent_ratio = round((colon_sent_count / max(colon_den_sents, 1)) * 100.0, 1)
    return {
        "prose_para_count": len(prose_paras),
        "total_words": sum(word_counts),
        "total_sents": len(sent_word_lens),
        "avg_sents_per_para": round(avg_sents_per_para, 2),
        "avg_words_per_para": round(avg_words_per_para, 1),
        "min_words_para": min(word_counts),
        "max_words_para": max(word_counts),
        "short_2_3_sent_para_ratio": round(short_1_3_ratio * 100, 1),
        "uniform_4_5_sent_para_ratio": round(uniform_4_5_ratio * 100, 1),
        "takeaway_end_count": takeaway_end_count,
        "takeaway_end_ratio": round(takeaway_ratio, 1),
        "avg_words_per_sent": round(avg_words_per_sent, 1),
        "median_words_per_sent": median_words_per_sent,
        "max_words_sent": max(sent_word_lens),
        "std_words_sent": round(std_words_sent, 1),
        "cv_sent_len": cv_sent_len,
        "over_45_words_count": over_45_count,
        "over_45_words_ratio": round(over_45_count / len(sent_word_lens) * 100, 1),
        "over_65_words_count": over_65_count,
        "staccato_count": staccato_count,
        "uniform_sent_run_count": uniform_sent_run_count,
        "colon_sent_count": colon_sent_count,
        "colon_den_sents": colon_den_sents,
        "colon_sent_ratio": colon_sent_ratio,
    }


_FACT_URL_RE = re.compile(r"https?://[^\s)>\]]+")
_FACT_CODE_RE = re.compile(r"`([^`\n]+)`")
_FACT_NUM_RE = re.compile(r"\bv?\d+(?:[.,:/\-–]\d+)*%?")
_FACT_DATE_RE = re.compile(r"(\d{1,2})/(\d{1,2})(?:/(\d{4}))?")
_FACT_LIST_RE = re.compile(r"(?m)^[\s#>*_-]*\d{1,2}[.)](?=\s)")
_FACT_SECTION_NUM_RE = re.compile(r"(?:^|\b)(?:chương|mục|hình|bảng|phụ\s+lục)\s+\d+(?:\.\d+)*\b", re.I)
_FACT_NUM_WORDS = {
    "hai": "2", "ba": "3", "bốn": "4", "năm": "5",
    "sáu": "6", "bảy": "7", "tám": "8", "chín": "9", "mười": "10",
}
_FACT_QTY_RE = re.compile(
    r"(?<!thứ )(?<!tháng )(?<!mùng )(?<!mồng )(?<!mười )\b(hai|ba|bốn|năm|sáu|bảy|tám|chín|mười)\s+"
    r"(?!\d|nay\b|ngoái\b|trước\b|sau\b|tới\b|qua\b|đó\b|này\b|kia\b|mới\b|học\b|mẹ\b|má\b|tôi\b|em\b|"
    r"con\b|ơi\b|là\b|mươi\b|trăm\b|nghìn\b|ngàn\b|triệu\b|tỷ\b|tỉ\b|"
    r"hai\b|ba\b|bốn\b|năm\b|sáu\b|bảy\b|tám\b|chín\b|mười\b)[^\W\d_]+"
)
_FACT_SHADES = {
    "chưa": r"\bchưa\b",
    "vẫn": r"\bvẫn\b",
    "chỉ": r"\bchỉ\b(?!\s+(?:ra|đạo|số|tiêu|định|dẫn|thị|huy|trích|tiết)\b)",
    "có thể": r"\bcó\s+thể\b",
    "có lẽ": r"\bcó\s+lẽ\b",
    "khoảng": r"\b(?:khoảng|chừng|tầm|gần)\s+\d",
    "ít nhất": r"\b(?:ít\s+nhất|tối\s+thiểu)\b",
    "tối đa": r"\b(?:tối\s+đa|nhiều\s+nhất|không\s+quá)\b",
    "dự kiến": r"\bdự\s+kiến\b",
    "hầu hết": r"\b(?:hầu\s+hết|phần\s+lớn|đa\s+số)\b",
}


def _canon_fact_number(num_str: str) -> str:
    n = num_str.replace("–", "-")
    core = n.rstrip("%")
    m_date = _FACT_DATE_RE.fullmatch(core)
    if m_date and 1 <= int(m_date.group(1)) <= 31 and 1 <= int(m_date.group(2)) <= 12:
        return "/".join(str(int(x)) for x in m_date.groups() if x)
    if re.fullmatch(r"\d{1,3}(?:[.,]\d{3})+", core):
        core = re.sub(r"[.,]", "", core)
    elif re.fullmatch(r"\d+,\d+", core):
        core = core.replace(",", ".")
    return core + n[len(n.rstrip("%")):]


def _date_without_year(num_str: str, pool: set) -> bool:
    m = _FACT_DATE_RE.fullmatch(num_str)
    return bool(m and num_str.count("/") == 1 and any(p.startswith(num_str + "/") for p in pool))


def _extract_facts_vi(text: str) -> dict:
    """Trích xuất số liệu chuẩn hóa, URL và từ giữ nghĩa P4 từ văn bản tiếng Việt."""
    cleaned_no_sec = _FACT_SECTION_NUM_RE.sub(" ", text)
    no_code = _FACT_LIST_RE.sub("", _FACT_CODE_RE.sub(" ", _FACT_URL_RE.sub(" ", cleaned_no_sec)))
    low = no_code.lower()
    numbers = {}
    for raw_n in _FACT_NUM_RE.findall(no_code):
        raw_n = raw_n.rstrip(".,")
        numbers.setdefault(_canon_fact_number(raw_n), raw_n)
    return {
        "_low": low,
        "_words": len(low.split()),
        "_numbers": numbers,
        "_quantities": {_FACT_NUM_WORDS[m.group(1)]: m.group(1) for m in _FACT_QTY_RE.finditer(low)},
        "số": set(numbers),
        "link": {u.rstrip(".,;:") for u in _FACT_URL_RE.findall(text)},
        "shades": {k: len(re.findall(rx, low)) for k, rx in _FACT_SHADES.items()},
    }


def _compare_facts_vi(src_text: str, dst_text: str, check_drops: bool = None) -> list:
    """
    Đối chiếu dữ kiện giữa tài liệu gốc/bản trước (src_text) và văn bản đích (dst_text) (nguyên tắc P1/P4).
    - Nếu check_drops=False (khi src là tài liệu tham khảo rất lớn so với 1 chương), ưu tiên cảnh báo
      số liệu/URL tự phát sinh mới không có trong nguồn.
    - Nếu check_drops=True (hoặc độ dài dst >= 50% src), kiểm cả số liệu bị rớt và từ giữ nghĩa ('có thể', 'vẫn', 'chưa'...) bị giảm.
    """
    src_f = _extract_facts_vi(src_text)
    dst_f = _extract_facts_vi(dst_text)
    if check_drops is None:
        check_drops = (
            src_f["_words"] >= 40
            and int(src_f["_words"] * 0.75) <= dst_f["_words"] <= int(src_f["_words"] * 1.35)
        )

    warnings = []
    added_nums = sorted(
        dst_f["_numbers"][n]
        for n in (dst_f["số"] - src_f["số"])
        if n not in src_f["_quantities"] and not _date_without_year(n, src_f["số"]) and len(n) >= 2
    )
    # Bỏ qua đánh số tiểu mục kiểu 1.1, 2.3.1 khi kiểm mới phát sinh
    added_nums = [x for x in added_nums if not re.fullmatch(r"\d+\.\d+(?:\.\d+)*", x)]
    if added_nums:
        warnings.append(
            f"[Đối chiếu dữ kiện P1/P4] Phát hiện {len(added_nums)} số liệu không thấy trong tài liệu nguồn: {', '.join(added_nums[:12])}"
        )

    added_links = sorted(dst_f["link"] - src_f["link"])
    if added_links:
        warnings.append(f"[Đối chiếu dữ kiện P1] Có URL mới chưa thấy trong nguồn: {', '.join(added_links[:5])}")

    if check_drops:
        lost_nums = sorted(
            src_f["_numbers"][n]
            for n in (src_f["số"] - dst_f["số"])
            if n not in dst_f["_quantities"] and not _date_without_year(n, dst_f["số"]) and not re.fullmatch(r"\d+\.\d+(?:\.\d+)*", src_f["_numbers"][n])
        )
        if lost_nums:
            warnings.append(f"[Đối chiếu dữ kiện P4] Thiếu {len(lost_nums)} số liệu so với bản gốc: {', '.join(lost_nums[:12])}")
        lost_links = sorted(src_f["link"] - dst_f["link"])
        if lost_links:
            warnings.append(f"[Đối chiếu dữ kiện] Bị mất URL so với bản gốc: {', '.join(lost_links[:5])}")
        fewer_shades = [
            f"{k} ({src_f['shades'][k]}->{dst_f['shades'][k]})"
            for k in src_f["shades"]
            if dst_f["shades"][k] < src_f["shades"][k]
        ]
        if len(fewer_shades) >= 2 or any("có thể" in x or "chưa" in x or "tối đa" in x for x in fewer_shades):
            warnings.append(
                f"[Đối chiếu mức khẳng định P4] Từ giữ nghĩa giảm so với nguồn ({', '.join(fewer_shades)}) -> rà lại kẻo đổi nghĩa gốc"
            )
    return warnings


def _compute_surface_ai_score(total_words: int, pattern_hits: int, doc_points: int, cv_sent_len) -> tuple:
    """
    Thang điểm dấu hiệu bề mặt 0-100 (tương tự viet-nhu-nguoi scan.py):
    kết hợp mật độ lỗi cụm AI/văn dịch, độ đều nhịp câu (CV < 0.35) và lỗi cấu trúc toàn bài.
    """
    density = (pattern_hits * 2.0) * 100.0 / max(total_words, 50)
    burst_penalty = max(0.0, 0.35 - cv_sent_len) * 40.0 if cv_sent_len is not None else 0.0
    score_val = min(100, round(density * 1.8 + burst_penalty + doc_points))
    if score_val < 10:
        band = "Sạch"
    elif score_val < 30:
        band = "Vài dấu hiệu"
    elif score_val < 60:
        band = "Nhiều dấu hiệu"
    else:
        band = "Dày đặc"
    return score_val, band


def inspect_docx_format(docx_path):
    """Kiểm tra mức độ tuân thủ format chuẩn ĐATN KMA (A4, lề, chân trang, font Times New Roman 14pt)."""
    report = {
        "sections": [],
        "normal_font": None,
        "normal_size_pt": None,
        "non_tnr_body_runs": 0,
        "wrong_sz_body_runs": 0,
        "issues": [],
    }
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ignored_fonts = {"Times New Roman", "Courier New", "Symbol", "Wingdings", "Webdings", "Cambria Math"}
    body_cover_done = False

    with zipfile.ZipFile(docx_path, "r") as z:
        doc_root = ET.fromstring(z.read("word/document.xml"))
        styles_root = ET.fromstring(z.read("word/styles.xml"))
        rels_map = {}
        if "word/_rels/document.xml.rels" in z.namelist():
            rels_root = ET.fromstring(z.read("word/_rels/document.xml.rels"))
            rels_map = {el.attrib["Id"]: el.attrib["Target"] for el in rels_root}

        # Kiểm tra style Normal
        for st in styles_root.findall("w:style", NS):
            if st.attrib.get(W_ID) == "Normal":
                rpr = st.find("w:rPr", NS)
                if rpr is not None:
                    rf = rpr.find("w:rFonts", NS)
                    if rf is not None:
                        report["normal_font"] = rf.attrib.get(f"{{{W_URI}}}ascii") or rf.attrib.get(f"{{{W_URI}}}hAnsi")
                    sz = rpr.find("w:sz", NS)
                    if sz is not None:
                        report["normal_size_pt"] = int(sz.attrib.get(W_VAL, "0")) / 2

        if report["normal_font"] and report["normal_font"] != "Times New Roman":
            report["issues"].append(f"Style Normal đang dùng font '{report['normal_font']}' (Chuẩn ĐATN: Times New Roman)")
        if report["normal_size_pt"] and report["normal_size_pt"] != 14:
            report["issues"].append(f"Style Normal đang dùng cỡ {report['normal_size_pt']}pt (Chuẩn ĐATN: 14pt)")

        body = doc_root.find("w:body", NS)
        if body is None:
            return report

        # Phân tích từng section
        sect_prs = body.findall(".//w:sectPr", NS)
        for idx, sp in enumerate(sect_prs, 1):
            pgsz = sp.find("w:pgSz", NS)
            pgmar = sp.find("w:pgMar", NS)
            pgnum = sp.find("w:pgNumType", NS)
            title_pg = sp.find("w:titlePg", NS) is not None
            frefs = [(fr.attrib.get(f"{{{W_URI}}}type", "default"), rels_map.get(fr.attrib.get(f"{{{R_NS}}}id"), "")) for fr in sp.findall("w:footerReference", NS)]

            w_val = pgsz.attrib.get(f"{{{W_URI}}}w", "0") if pgsz is not None else "0"
            h_val = pgsz.attrib.get(f"{{{W_URI}}}h", "0") if pgsz is not None else "0"
            is_a4 = (abs(int(w_val) - 11906) <= 30 and abs(int(h_val) - 16838) <= 30)

            m = {k.split("}")[-1]: int(v) for k, v in pgmar.attrib.items()} if pgmar is not None else {}
            margins_ok = (
                abs(m.get("top", 0) - 1134) <= 30
                and abs(m.get("bottom", 0) - 1134) <= 30
                and abs(m.get("left", 0) - 1701) <= 30
                and abs(m.get("right", 0) - 851) <= 30
                and abs(m.get("footer", 0) - 709) <= 40
            )

            has_page_field = False
            footer_centered = True
            footer_texts = []
            for ftype, target in frefs:
                part = f"word/{target}" if target and not target.startswith("word/") else target
                if part and part in z.namelist():
                    froot = ET.fromstring(z.read(part))
                    instrs = "".join(t.text for t in froot.findall(".//w:instrText", NS) if t.text)
                    txts = "".join(t.text for t in froot.findall(".//w:t", NS) if t.text).strip()
                    if "PAGE" in instrs:
                        has_page_field = True
                        jc_el = froot.find(".//w:jc", NS)
                        if jc_el is not None and jc_el.attrib.get(W_VAL) != "center":
                            footer_centered = False
                    elif txts:
                        footer_texts.append((ftype, txts))

            sec_info = {
                "index": idx,
                "is_a4": is_a4,
                "w_cm": round(int(w_val) / 567, 1),
                "h_cm": round(int(h_val) / 567, 1),
                "top_cm": round(m.get("top", 0) / 567, 2),
                "bottom_cm": round(m.get("bottom", 0) / 567, 2),
                "left_cm": round(m.get("left", 0) / 567, 2),
                "right_cm": round(m.get("right", 0) / 567, 2),
                "footer_cm": round(m.get("footer", 0) / 567, 2),
                "margins_ok": margins_ok,
                "has_page_field": has_page_field,
                "footer_centered": footer_centered,
                "pg_fmt": pgnum.attrib.get(f"{{{W_URI}}}fmt", "decimal") if pgnum is not None else "decimal",
                "title_pg": title_pg,
                "hardcoded_footers": footer_texts,
            }
            report["sections"].append(sec_info)

            if not is_a4:
                report["issues"].append(f"[Format Section {idx}] Khổ giấy hiện là {sec_info['w_cm']}x{sec_info['h_cm']}cm (Chuẩn ĐATN: A4 21.0x29.7cm)")
            if not margins_ok:
                report["issues"].append(
                    f"[Format Section {idx}] Căn lề chưa chuẩn ({sec_info['top_cm']}-{sec_info['bottom_cm']}-{sec_info['left_cm']}-{sec_info['right_cm']}cm, "
                    f"footer={sec_info['footer_cm']}cm; Chuẩn ĐATN: 2.0-2.0-3.0-1.5cm, footer=1.25cm)"
                )
            if footer_texts:
                report["issues"].append(f"[Format Section {idx}] Chân trang đang gõ cứng số/chữ '{footer_texts[0][1]}' thay vì dùng trường PAGE tự động")
            if idx == len(sect_prs) and not has_page_field:
                report["issues"].append(f"[Format Section {idx}] Chân trang phần nội dung chưa có trường đánh số trang tự động (PAGE)")
            if has_page_field and not footer_centered:
                report["issues"].append(f"[Format Section {idx}] Số trang ở chân trang chưa căn giữa")

        # Đếm font lệch trong phần thân (bỏ qua phần trang bìa và style Code)
        for p in body.findall(".//w:p", NS):
            ppr = p.find("w:pPr", NS)
            st = "Normal"
            if ppr is not None:
                ps = ppr.find("w:pStyle", NS)
                if ps is not None:
                    st = ps.attrib.get(W_VAL, "Normal")
            txt = "".join(t.text for t in p.findall(".//w:t", NS) if t.text).strip()
            low_txt = txt.lower()
            if st in ("UH1", "Heading1", "Mclc") or any(k in low_txt for k in ("mục lục", "lời nói đầu", "lời mở đầu", "chương 1", "chương i")):
                body_cover_done = True
            if not body_cover_done or st in ("TrangBa", "Code"):
                continue
            for r in p.findall("w:r", NS):
                rpr = r.find("w:rPr", NS)
                if rpr is None:
                    continue
                if rpr.find("w:vertAlign", NS) is not None:
                    continue
                rf = rpr.find("w:rFonts", NS)
                if rf is not None:
                    f_name = rf.attrib.get(f"{{{W_URI}}}ascii") or rf.attrib.get(f"{{{W_URI}}}hAnsi") or ""
                    if f_name and f_name not in ignored_fonts:
                        report["non_tnr_body_runs"] += 1
                sz = rpr.find("w:sz", NS)
                if sz is not None and st in ("Normal", "Tiliuthamkho", "Tnhnhv", "Tnbng"):
                    if sz.attrib.get(W_VAL) not in ("28", ""):
                        report["wrong_sz_body_runs"] += 1

        if report["non_tnr_body_runs"] > 0:
            report["issues"].append(
                f"[Format Font] Phát hiện {report['non_tnr_body_runs']} đoạn/cụm chữ trong thân bài dùng font khác Times New Roman"
            )
        if report["wrong_sz_body_runs"] > 0:
            report["issues"].append(
                f"[Format Cỡ chữ] Phát hiện {report['wrong_sz_body_runs']} đoạn/cụm chữ trong thân bài khác cỡ 14pt"
            )

    return report


# ==============================================================================
# 3. CÁC LỆNH ĐỌC & KIỂM TRA: inspect, check, score-local
# ==============================================================================
def cmd_inspect(docx_path):
    if not os.path.exists(docx_path):
        print(f"[LỖI] Không tìm thấy file: {docx_path}")
        sys.exit(1)
    paragraphs, _ = extract_docx_paragraphs(docx_path)
    fmt = inspect_docx_format(docx_path)
    print(f"=== KIỂM TRA KHUNG FILE BÁO CÁO / ĐỒ ÁN: {docx_path} ===")
    print(f"Tổng số đoạn văn / phần tử: {len(paragraphs)}")
    if fmt["sections"]:
        s0 = fmt["sections"][-1]
        print(
            f"Khổ giấy (thân bài): {s0['w_cm']}x{s0['h_cm']}cm ({'A4 Đạt' if s0['is_a4'] else 'Chưa phải A4 21.0x29.7cm'})"
        )
        print(
            f"Căn lề (thân bài): Trên={s0['top_cm']}cm, Dưới={s0['bottom_cm']}cm, Trái={s0['left_cm']}cm, Phải={s0['right_cm']}cm, "
            f"Footer={s0['footer_cm']}cm (Chuẩn KMA: 2.0 - 2.0 - 3.0 - 1.5 cm, Footer: 1.25 cm)"
        )
        sec_summaries = []
        for s in fmt["sections"]:
            pg_desc = "Không số" if not s["has_page_field"] else (f"La Mã ({s['pg_fmt']})" if s["pg_fmt"] == "lowerRoman" else "Ả Rập")
            if s["hardcoded_footers"]:
                pg_desc += " [GÕ CỨNG]"
            sec_summaries.append(f"S{s['index']}:{pg_desc}")
        print(f"Chân trang theo Section ({len(fmt['sections'])} section): " + " | ".join(sec_summaries))
    if fmt["issues"]:
        print(f"Cảnh báo Format chưa đồng bộ chuẩn ĐATN ({len(fmt['issues'])} điểm — có thể sửa tự động bằng lệnh 'apply-format'):")
        for fi in fmt["issues"]:
            print("  !", fi)

    placeholders = [
        (idx, st, txt) for idx, st, txt, _ in paragraphs if PLACEHOLDER_MARKER_RE.search(txt)
    ]
    print(f"Số khung giữ chỗ hình ảnh [TODO tự chèn ảnh]: {len(placeholders)}")

    stats = compute_videtect_stats([txt for _, st, txt, _ in paragraphs if st in ("Normal", "BodyText")])
    if stats:
        print(
            f"Thống kê cấu trúc đoạn (chuẩn ViDetect 2405.03206v1): "
            f"{stats['prose_para_count']} đoạn văn xuôi | TB {stats['avg_sents_per_para']} câu/đoạn "
            f"({stats['avg_words_per_para']} từ/đoạn, min={stats['min_words_para']}, max={stats['max_words_para']}) | "
            f"Nhịp câu TB {stats['avg_words_per_sent']} từ (CV={stats['cv_sent_len']})"
        )

    print("\n--- CÂY ĐỀ MỤC, HÌNH VẼ & BẢNG BIỂU TRONG FILE ---")
    target_styles = {
        "TrangBa", "UH1", "UH2", "Heading1", "Heading2", "Heading3", "Heading4",
        "Tnbng", "Tnhnhv", "Hnhv", "Mclc",
    }
    for idx, st, txt, has_img in paragraphs:
        if st in target_styles or PLACEHOLDER_MARKER_RE.search(txt):
            img_tag = " [CÓ ẢNH/OBJECT]" if has_img else ""
            print(f"[Đoạn {idx:4d}] [{st:10s}]{img_tag} {txt[:120]}")


def load_items_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    items = []
    raw_lines = []
    if ext == ".docx":
        paragraphs, _ = extract_docx_paragraphs(file_path)
        for idx, st, txt, has_img in paragraphs:
            if txt:
                items.append((f"Đoạn {idx}", st, txt, has_img))
    elif ext == ".pdf":
        try:
            import pymupdf
            doc = pymupdf.open(file_path)
            for pno, page in enumerate(doc, 1):
                blocks = page.get_text("blocks")
                block_texts = [b[4].strip() for b in blocks if len(b) >= 5 and isinstance(b[4], str) and b[4].strip()]
                for para in normalize_vietnamese_paragraphs("\n".join(block_texts)):
                    items.append((f"Trang {pno}", "Normal", para, False))
        except ImportError:
            print("[LỖI] Cần cài pymupdf để đọc trực tiếp file PDF.")
            sys.exit(1)
    else:
        with open(file_path, "r", encoding="utf-8-sig") as f:
            raw_lines = f.readlines()
        in_code = False
        for i, line in enumerate(raw_lines, 1):
            s = line.strip()
            if s.startswith("```"):
                if s.lower().startswith("```mermaid"):
                    items.append((f"Dòng {i}", "Mermaid", s, False))
                in_code = not in_code
                continue
            if in_code or not s:
                continue
            st = "Heading" if s.startswith("#") else "Normal"
            items.append((f"Dòng {i}", st, s, False))
    return items, raw_lines


def _scan_prose_report(items, mode="do-an", ext=".md", no_cite=True, require_cite=False):
    """Phân tích tổng hợp lỗi AI, văn dịch sâu, dấu câu, nhịp câu CV và điểm dấu hiệu bề mặt 0-100."""
    if require_cite:
        no_cite = False
    mode = mode.lower()
    issues = []
    todos_images = []
    em_dash_count = 0
    md_bold_bullet_count = 0
    backtick_count = 0
    raw_dash_bullet_count = 0
    code_explain_heading_count = 0
    in_refs_section = False
    in_summary_section = False
    in_chap_conclusion = False
    in_appendix_section = False
    summary_words = 0
    refs_entry_count = 0
    in_body_citations = 0
    en_paren_total = 0
    pattern_hit_count = 0
    false_contrast_count = 0
    self_explain_count = 0
    this_it_open_count = 0
    thing_subj_count = 0
    passive_calque_count = 0
    prose_all_texts = []
    connector_counts = {
        "Ngoài ra": 0,
        "Bên cạnh đó": 0,
        "Đồng thời": 0,
        "Song song": 0,
        "Tiếp theo": 0,
        "Trước hết": 0,
        "Nhóm thứ nhất": 0,
    }
    discourse_group_counts = {
        "Chuyển ý đối lập/nhượng bộ (Trái lại / Tuy vậy / Dù vậy)": 0,
        "Chuyển ý nhân quả (Vì vậy / Vì thế / Bởi vậy / Bởi thế / Do đó)": 0,
        "Mở đoạn định vị lặp (Trên nền tảng / Trên máy chủ / Trong thực tế / Trong thực chiến)": 0,
        "Mở đoạn gán nhãn đánh giá (Điểm khó đối phó / Một rủi ro / Hạn chế lớn nhất / Giá trị thực chất)": 0,
    }

    patterns = list(AI_TELL_PATTERNS_VI)
    if mode == "do-an":
        patterns.extend(DO_AN_ONLY_PATTERNS)

    false_contrast_all_re = re.compile(
        r"(?i)(?:\bkhông\s+(?:chỉ|đơn\s+thuần|phải)\b.{0,80}?\bmà\s+(?:còn|là)\b|"
        r"\bkhông\s+(?!chỉ\b|đơn\s+thuần\b|phải\b)[^.!?\n,;:]{1,60},\s*mà\s+(?!còn\b)|"
        r"\bkhông\s+(?:còn\s+)?(?:là|ở|nằm\s+ở)\s+[^.!?\n]{1,80}\bmà\s+(?:là|ở|nằm\s+ở)\b|"
        r"(?:^|[.!?]\s+)(?:đừng|thay\s+vì)\s+[^.!?\n,]{2,40},\s*hãy\b|"
        r"\bchứ\s+không(?!\s+thì)\b|"
        r"\blà\s+[^.!?\n,]{2,40},\s*không\s+phải\s+(?:là\s+)?(?!làm|gõ|cài|cấu\s+hình|chờ|lo|trả|đi|mua|sửa|viết|đợi|nghĩ|hỏi|tốn|mất|ai|lúc|khi)\w)"
    )
    self_explain_re = re.compile(
        r"(?i)(?:^|(?<=[.!?]\s))(?:tức\s+là|nghĩa\s+là|có\s+nghĩa\s+là|(?:hay\s+)?nói\s+(?:một\s+)?cách\s+khác)\b"
    )
    this_it_open_re = re.compile(r"(?i)(?:^|(?<=[.!?]\s))(?:điều\s+này|điều\s+đó|nó)\b")
    you_pronoun_re = re.compile(
        r"(?i)\b(?:cho\s+phép\s+bạn|giúp\s+bạn|của\s+bạn|bạn\s+có\s+thể|bạn\s+nên|bạn\s+cần|hãy\s+chắc\s+chắn\s+rằng\s+bạn)\b"
    )

    for loc, st, txt, _has_img in items:
        if PLACEHOLDER_MARKER_RE.search(txt):
            todos_images.append(f"[{loc}] {txt[:115]}")
            if re.search(r"(?i)(loại\s+hình\s+gợi\s+ý|gợi\s+ý\s+nội\s+dung\s+cần\s+có|\.(?:pdf|md|html)\b)", txt):
                issues.append(
                    f"[{loc}] Dấu vết quy trình sinh nội dung ở khung giữ chỗ hình (còn 'Gợi ý nội dung' hoặc tên file cục bộ): \"{txt[:100]}...\""
                )
            continue

        if st == "Mermaid":
            issues.append(
                f"[{loc}] Vi phạm quy tắc Hình ảnh: Phát hiện khối sơ đồ ```mermaid (phải để trống khung giữ chỗ hình)"
            )
            continue
        if ASCII_DIAGRAM_RE.search(txt):
            issues.append(
                f"[{loc}] Vi phạm quy tắc Hình ảnh: Phát hiện sơ đồ ASCII '{txt[:60]}...' (phải thay bằng khung giữ chỗ hình ảnh)"
            )

        if st in ("Heading", "Heading1", "Heading2", "Heading3", "Heading4", "UH1", "UH2", "Mclc"):
            clean_h = re.sub(r"^#+\s*", "", txt).strip()
            h_low = clean_h.lower()
            in_refs_section = "tài liệu tham khảo" in h_low
            in_summary_section = "tóm tắt" in h_low
            in_chap_conclusion = bool(re.search(r"kết\s+luận\s+chương", h_low))
            if "phụ lục" in h_low:
                in_appendix_section = True
            elif re.match(r"^(?:chương|kết\s+luận|tài\s+liệu\s+tham\s+khảo)\b", h_low):
                in_appendix_section = False
            if clean_h.endswith((".", ":", "?", ";")):
                issues.append(f"[{loc}] Lỗi đề mục (có dấu câu ở cuối): '{clean_h}'")
            elif "?" in clean_h or re.search(r"(?i)\b(là\s+gì|như\s+thế\s+nào|tại\s+sao|vì\s+sao)\b", clean_h):
                issues.append(f"[{loc}] Lỗi đề mục (đề mục dạng câu hỏi nghi vấn): '{clean_h}'")
            continue

        if in_refs_section:
            if re.match(r"^\[\d+\]", txt.strip()):
                refs_entry_count += 1
            continue

        if in_summary_section and st in ("Normal", "NoIndent"):
            summary_words += count_words_vi(txt)

        if in_chap_conclusion and st in ("Normal", "NoIndent"):
            wc_concl = count_words_vi(txt)
            comma_cnt = txt.count(",")
            if wc_concl <= 165 and comma_cnt >= 7:
                issues.append(
                    f"[{loc}] [Kết luận chương nhồi từ khóa] Đoạn Kết luận chương ({wc_concl} từ) chứa {comma_cnt} dấu phẩy liệt kê dày đặc thuật ngữ kiểu bản tóm tắt AI -> cần viết thành nhận định đúc kết của sinh viên và lý do chọn công cụ cho chương sau"
                )

        if "—" in txt:
            em_dash_count += txt.count("—")
        if re.match(r"^[-*+•]\s*\*\*[^*]{2,50}\*\*\s*:", txt):
            md_bold_bullet_count += 1
        if ext == ".docx" and st != "Code":
            if "`" in txt:
                backtick_count += txt.count("`")
            if re.match(r"^-\s+\S", txt):
                raw_dash_bullet_count += 1
        if re.match(r"(?i)^giải\s+thích\s+(lệnh|đoạn\s+mã)\s*:", txt.strip()):
            code_explain_heading_count += 1

        if st in ("Normal", "NoIndent") and not st.startswith("TOC"):
            prose_all_texts.append(txt)
            in_body_citations += len(re.findall(r"\[\d+(?:\s*[-,]\s*\d+)*\]", txt))
            en_paren_total += len(re.findall(r"\([A-Za-z][A-Za-z0-9\s/\-\.,:&]{2,}\)", txt))
            for conn in connector_counts:
                connector_counts[conn] += len(re.findall(r"(?i)\b" + re.escape(conn) + r"\b", txt))
            discourse_group_counts["Chuyển ý đối lập/nhượng bộ (Trái lại / Tuy vậy / Dù vậy)"] += len(
                re.findall(r"(?i)(?:^|(?<=[.!?]\s))(?:trái\s+lại|tuy\s+vậy|dù\s+vậy)\b", txt)
            )
            discourse_group_counts["Chuyển ý nhân quả (Vì vậy / Vì thế / Bởi vậy / Bởi thế / Do đó)"] += len(
                re.findall(r"(?i)(?:^|(?<=[.!?]\s))(?:vì\s+vậy|vì\s+thế|bởi\s+vậy|bởi\s+thế|do\s+đó)\b", txt)
            )
            discourse_group_counts["Mở đoạn định vị lặp (Trên nền tảng / Trên máy chủ / Trong thực tế / Trong thực chiến)"] += len(
                re.findall(r"(?i)^(?:trên\s+(?:các\s+)?nền\s+tảng|trên\s+(?:các\s+)?máy\s+chủ|trong\s+thực\s+(?:tế|chiến))\b", txt.strip())
            )
            discourse_group_counts["Mở đoạn gán nhãn đánh giá (Điểm khó đối phó / Một rủi ro / Hạn chế lớn nhất / Giá trị thực chất)"] += len(
                re.findall(
                    r"(?i)(?:^|(?<=[.!?]\s))(?:điểm\s+khó\s+đối\s+phó|một\s+rủi\s+ro\s+(?:quan\s+trọng|thực\s+tế)|một\s+điểm\s+yếu|hạn\s+chế\s+lớn\s+nhất|giá\s+trị\s+thực\s+chất\s+của)\b",
                    txt.strip(),
                )
            )
            false_contrast_count += len(false_contrast_all_re.findall(txt))
            self_explain_count += len(self_explain_re.findall(txt))
            this_it_open_count += len(this_it_open_re.findall(txt.strip()))
            thing_subj_count += len(THING_SUBJECT_RE.findall(txt))
            passive_calque_count += len(PASSIVE_CALQUE_RE.findall(txt))
            if not in_appendix_section and you_pronoun_re.search(txt):
                pattern_hit_count += 1
                issues.append(
                    f"[{loc}] Lỗi xưng hô lạc vai (viet-nhu-nguoi Mục 36/56): Không xưng 'bạn / của bạn / giúp bạn' trong thân bài BTL/ĐATN -> \"{txt[:105]}...\""
                )

        if st != "Code" and not st.startswith("TOC"):
            for pat, msg in patterns:
                if re.search(pat, txt):
                    pattern_hit_count += 1
                    issues.append(f"[{loc}] {msg} -> \"{txt[:110]}...\"")

    doc_points = 0
    if em_dash_count > 0:
        doc_points += 3
    if backtick_count > 0:
        doc_points += 4
        issues.append(
            f"[Dấu vết Markdown] Phát hiện {backtick_count} ký tự backtick (`) sót trong văn xuôi/bảng .docx -> phải gỡ sạch khỏi Word"
        )
    if raw_dash_bullet_count > 0:
        doc_points += 3
        issues.append(
            f"[Dấu vết Markdown] Phát hiện {raw_dash_bullet_count} dòng bắt đầu bằng '- ' thô trong .docx -> chuyển sang văn xuôi hoặc bullet Word chuẩn"
        )
    if code_explain_heading_count >= 3:
        doc_points += 4
        issues.append(
            f"[Khuôn lặp AI] Nhãn 'Giải thích lệnh / đoạn mã:' lặp máy móc {code_explain_heading_count} lần -> nên diễn giải trực tiếp thành đoạn văn tự nhiên dưới khối mã"
        )
    if summary_words > 260:
        issues.append(
            f"[Tóm tắt quá dài] Phần Tóm tắt dài {summary_words} từ (> 260 từ) -> nên rút gọn cô đọng mục tiêu/phạm vi khoảng 150-230 từ, không chép số liệu chi tiết thân bài"
        )
    if refs_entry_count >= 3 and in_body_citations == 0 and not no_cite:
        issues.append(
            f"[Liêm chính học thuật] Danh mục Tài liệu tham khảo có {refs_entry_count} nguồn nhưng thân bài không có trích dẫn [1], [2]... nào -> cần gắn trích dẫn đúng vị trí định nghĩa/số liệu/bảng biểu"
        )
    for conn, cnt in connector_counts.items():
        if cnt >= 5:
            doc_points += 3
            issues.append(
                f"[Từ nối lặp khuôn] Cụm '{conn}' lặp {cnt} lần trong văn xuôi (nên giữ <= 3-4 lần và đa dạng cách chuyển ý)"
            )
    discourse_thresholds = {
        "Chuyển ý đối lập/nhượng bộ (Trái lại / Tuy vậy / Dù vậy)": 4,
        "Chuyển ý nhân quả (Vì vậy / Vì thế / Bởi vậy / Bởi thế / Do đó)": 5,
        "Mở đoạn định vị lặp (Trên nền tảng / Trên máy chủ / Trong thực tế / Trong thực chiến)": 4,
        "Mở đoạn gán nhãn đánh giá (Điểm khó đối phó / Một rủi ro / Hạn chế lớn nhất / Giá trị thực chất)": 3,
    }
    for grp_name, g_cnt in discourse_group_counts.items():
        if g_cnt >= discourse_thresholds[grp_name]:
            doc_points += 4
            issues.append(
                f"[Lặp cụm chuyển ý/mở đoạn] Nhóm '{grp_name}' xuất hiện {g_cnt} lần -> cần chuyển ý tự nhiên bằng chủ ngữ kỹ thuật hoặc tách/gộp mạch câu"
            )

    if false_contrast_count >= 3:
        doc_points += 5
        issues.append(
            f"[Đối lập lặp cả bài (viet-nhu-nguoi Mục 46)] Cấu trúc 'không X mà Y / X chứ không Y' lặp {false_contrast_count} lần -> chỉ giữ tối đa 1-2 chỗ thực sự sửa hiểu lầm"
        )
    if self_explain_count >= 3:
        doc_points += 4
        issues.append(
            f"[Giải nghĩa hộ lặp khuôn (viet-nhu-nguoi Mục 48)] Cụm 'Tức là / Nghĩa là / Nói cách khác...' mở câu {self_explain_count} lần -> gộp ý thẳng vào câu trước"
        )
    if this_it_open_count >= 4:
        doc_points += 4
        issues.append(
            f"[Văn dịch 'Điều này/Nó' mở câu] Cụm 'Điều này / Điều đó / Nó' mở câu {this_it_open_count} lần -> thay bằng chủ ngữ kỹ thuật cụ thể"
        )
    if thing_subj_count >= 4:
        doc_points += 4
        issues.append(
            f"[Vật làm chủ ngữ lặp (viet-nhu-nguoi Mục 58)] Mẫu 'Hệ thống/Công cụ/Bảng... cho phép/mang lại/đảm bảo rằng' lặp {thing_subj_count} lần -> đổi sang câu chủ động"
        )

    prose_low_all = " ".join(prose_all_texts).lower()
    prose_total_words = len(prose_low_all.split())
    filler_total = sum(len(re.findall(fp, prose_low_all)) for fp in FILLER_WORDS_VI) if prose_total_words else 0
    filler_per_1k = round(filler_total * 1000.0 / max(prose_total_words, 1), 1)
    passive_per_1k = round(passive_calque_count * 1000.0 / max(prose_total_words, 1), 1)
    if prose_total_words >= 100 and filler_per_1k >= 115.0:
        doc_points += 10
        issues.append(
            f"[Chữ đệm dịch dày (viet-nhu-nguoi Nhóm K)] Mật độ chữ đệm dịch ({filler_per_1k}/1.000 từ: một, các, những, được, có thể, sẽ, đã, rằng, việc, sự...) cao bất thường (chuẩn báo cáo KMA ~67/1.000 từ)"
        )
    if prose_total_words >= 120 and passive_calque_count >= 5 and passive_per_1k >= 6.0:
        doc_points += 5
        issues.append(
            f"[Mật độ bị động dịch sát cao (viet-nhu-nguoi Mục 58)] Có {passive_calque_count} cụm 'được <động từ>' ({passive_per_1k}/1.000 từ) -> ưu tiên câu chủ động nêu rõ giao thức/thành phần xử lý"
        )

    stats = compute_videtect_stats(prose_all_texts)
    cv_val = stats["cv_sent_len"] if stats else None
    if stats:
        en_paren_per_100w = round(en_paren_total / max(stats["total_words"], 1) * 100, 2)
        if en_paren_total > 45 and en_paren_per_100w > 1.2:
            doc_points += 4
            issues.append(
                f"[Dịch thuật / Dày chú thích] Phát hiện {en_paren_total} cụm tiếng Anh trong ngoặc ({en_paren_per_100w}/100 từ) -> bỏ chú thích sau từ phổ thông, chỉ giữ 1 lần đầu cho thuật ngữ chuyên sâu"
            )
        if stats["avg_words_per_sent"] > 34.0 or stats["over_45_words_ratio"] > 25.0:
            doc_points += 4
            issues.append(
                f"[Nhịp câu quá dài/đều kiểu AI] TB {stats['avg_words_per_sent']} từ/câu (trung vị {stats['median_words_per_sent']}, "
                f"{stats['over_45_words_count']}/{stats['total_sents']} câu > 45 từ = {stats['over_45_words_ratio']}%, câu dài nhất {stats['max_words_sent']} từ). "
                "Cần tách câu đa dạng nhịp ngắn-vừa (~18-28 từ/câu)."
            )
        if stats["over_65_words_count"] > 0:
            doc_points += 3
            issues.append(
                f"[Câu quá dài] Có {stats['over_65_words_count']} câu dài hơn 65 từ (tối đa {stats['max_words_sent']} từ) -> tách thành 2-3 câu rõ ý"
            )
        if stats["total_sents"] >= 10 and stats["cv_sent_len"] < 0.30:
            doc_points += 4
            issues.append(
                f"[Nhịp câu đơn điệu (CV={stats['cv_sent_len']})] Độ dài các câu biến thiên quá đều (chuẩn tự nhiên thường CV >= 0.38)"
            )
        if stats["staccato_count"] > 0:
            doc_points += 5
            issues.append(
                f"[Chuỗi câu cụt (viet-nhu-nguoi Mục 54)] Có {stats['staccato_count']} chỗ ghép 3+ câu cực ngắn (<= 4 từ) liên tiếp -> nối lại theo mạch ý tự nhiên"
            )
        if stats["uniform_sent_run_count"] > 0:
            doc_points += 4
            issues.append(
                f"[Chuỗi câu dài đều nhau] Có {stats['uniform_sent_run_count']} đoạn chứa 3+ câu dài (>= 20 từ) liên tiếp gần bằng nhau về số từ"
            )
        if stats["colon_den_sents"] >= 15 and stats["colon_sent_count"] >= 6 and stats["colon_sent_ratio"] >= 15.0:
            doc_points += 6
            issues.append(
                f"[Lạm dụng dấu hai chấm ':' giữa câu (Mục 48)] Có {stats['colon_sent_count']}/{stats['colon_den_sents']} câu ({stats['colon_sent_ratio']}%) dùng dấu ':' không phải dẫn danh sách/lệnh"
            )
        if stats["prose_para_count"] >= 15 and stats["uniform_4_5_sent_para_ratio"] >= 50.0:
            doc_points += 6
            issues.append(
                f"[Cấu trúc đoạn đồng đều kiểu khuôn 4 bước] Có {stats['uniform_4_5_sent_para_ratio']}% số đoạn văn đóng cứng ở 4-5 câu (tỷ lệ đoạn ngắn 1-3 câu chỉ {stats['short_2_3_sent_para_ratio']}%). "
                "Cần xen kẽ đoạn dẫn ngắn 1-3 câu với đoạn phân tích sâu giống báo cáo sinh viên thực tế."
            )
        if stats["takeaway_end_count"] >= 6 or stats["takeaway_end_ratio"] >= 14.0:
            doc_points += 6
            issues.append(
                f"[Khuôn câu chốt đoạn lặp lại] Có {stats['takeaway_end_count']} đoạn ({stats['takeaway_end_ratio']}%) kết thúc bằng câu nhận xét/khuyên bảo khuôn mẫu ('giúp quản trị viên...', 'Vì thế/Vì vậy...'). "
                "Hãy kết đoạn ngay khi mô tả xong cơ chế kỹ thuật."
            )

    surf_score, surf_band = _compute_surface_ai_score(
        total_words=stats["total_words"] if stats else prose_total_words,
        pattern_hits=pattern_hit_count,
        doc_points=doc_points,
        cv_sent_len=cv_val,
    )
    return {
        "issues": issues,
        "todos_images": todos_images,
        "em_dash_count": em_dash_count,
        "md_bold_bullet_count": md_bold_bullet_count,
        "backtick_count": backtick_count,
        "in_body_citations": in_body_citations,
        "refs_entry_count": refs_entry_count,
        "en_paren_total": en_paren_total,
        "pattern_hit_count": pattern_hit_count,
        "filler_per_1k": filler_per_1k,
        "passive_per_1k": passive_per_1k,
        "stats": stats,
        "surf_score": surf_score,
        "surf_band": surf_band,
    }


def cmd_check(file_path, mode="do-an", no_cite=True, require_cite=False, source_path=None):
    if not os.path.exists(file_path):
        print(f"[LỖI] Không tìm thấy file: {file_path}")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    items, _ = load_items_from_file(file_path)
    res = _scan_prose_report(items, mode=mode, ext=ext, no_cite=no_cite, require_cite=require_cite)
    issues = res["issues"]

    if ext == ".docx":
        fmt = inspect_docx_format(file_path)
        for fi in fmt["issues"]:
            issues.append(fi)
        for xml_prob in validate_docx_ooxml(file_path):
            issues.append(xml_prob)

    if source_path:
        if not os.path.exists(source_path):
            issues.append(f"[LỖI] Không tìm thấy file nguồn đối chiếu: {source_path}")
        else:
            src_items, _ = load_items_from_file(source_path)
            src_prose = "\n".join(t for _, st, t, _ in src_items if st != "Code" and not PLACEHOLDER_MARKER_RE.search(t))
            dst_prose = "\n".join(t for _, st, t, _ in items if st != "Code" and not PLACEHOLDER_MARKER_RE.search(t))
            issues.extend(_compare_facts_vi(src_prose, dst_prose))

    stats = res["stats"]
    mode_label = "ĐỒ ÁN TỐT NGHIỆP (do-an)" if mode.lower() == "do-an" else "BÀI TẬP LỚN / CHUYÊN ĐỀ MÔN (btl)"
    print(f"=== KẾT QUẢ RÀ SOÁT [{mode_label}]: {file_path} ===")
    print(
        f"- Điểm dấu hiệu bề mặt (viet-nhu-nguoi 0-100): {res['surf_score']}/100 ({res['surf_band']}) "
        f"| Chữ đệm dịch: {res['filler_per_1k']}/1k từ | Bị động: {res['passive_per_1k']}/1k từ"
    )
    print(
        f"- Số dấu gạch ngang dài (—): {res['em_dash_count']} | Backtick (`) sót: {res['backtick_count']} "
        f"| Trích dẫn [n] trong thân bài: {res['in_body_citations']}/{res['refs_entry_count']} nguồn | Chú thích (EN): {res['en_paren_total']}"
    )
    if res["md_bold_bullet_count"] > 0:
        print(
            f"- Số dòng gạch đầu dòng dạng '- **Nhãn:**': {res['md_bold_bullet_count']} "
            "(Lưu ý: Trong 7.222 mẫu VietAIDetector, văn bản người viết hoàn toàn không lạm dụng '**' trang trí)"
        )
    if stats:
        print(
            f"- Phân bố đoạn văn xuôi: TB {stats['avg_sents_per_para']} câu/đoạn "
            f"({stats['avg_words_per_para']} từ/đoạn) | Đoạn 1-3 câu: {stats['short_2_3_sent_para_ratio']}% "
            f"| Đoạn 4-5 câu: {stats['uniform_4_5_sent_para_ratio']}% | Câu chốt đoạn khuôn mẫu: {stats['takeaway_end_count']} ({stats['takeaway_end_ratio']}%)"
        )
        print(
            f"- Nhịp độ dài câu: TB {stats['avg_words_per_sent']} từ/câu (trung vị {stats['median_words_per_sent']}, "
            f"CV={stats['cv_sent_len']}, max={stats['max_words_sent']}) | Câu >45 từ: {stats['over_45_words_count']} ({stats['over_45_words_ratio']}%) | Câu có ':' giữa câu: {stats['colon_sent_count']} ({stats['colon_sent_ratio']}%)"
        )

    if res["todos_images"]:
        print(f"- Danh sách [{len(res['todos_images'])}] vị trí giữ chỗ hình ảnh chờ bạn tự chèn ảnh thật:")
        for td in res["todos_images"][:25]:
            print("  [TODO ẢNH]", td)

    if not issues:
        print("- [ĐẠT] Không phát hiện lỗi đề mục, sơ đồ tự vẽ hay mẫu câu AI tiếng Việt trong danh sách kiểm tra!")
    else:
        print(f"- [CẦN SỬA] Phát hiện {len(issues)} điểm cần lưu ý/chỉnh sửa:")
        for iss in issues[:40]:
            print("  *", iss)
        if len(issues) > 40:
            print(f"  ... và {len(issues) - 40} điểm khác.")


def cmd_score_local(file_path):
    """
    Chuẩn hóa văn bản theo pipeline VietAIDetector (normalizer.py) + đo chỉ số ViDetect + CV +
    chỉ số bề mặt `viet-nhu-nguoi` (0-100), và nếu môi trường có cài torch + transformers thì
    chấm thử qua model cục bộ vietnamese-ai-detection-v2 (PhoBERT-large).
    """
    if not os.path.exists(file_path):
        print(f"[LỖI] Không tìm thấy file: {file_path}")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    items, _ = load_items_from_file(file_path)
    raw_text = "\n".join(
        txt for _, st, txt, _ in items
        if st in ("Normal", "NoIndent") and not PLACEHOLDER_MARKER_RE.search(txt)
    )
    norm_paras = normalize_vietnamese_paragraphs(raw_text)
    res = _scan_prose_report(items, mode="btl", ext=ext)
    print(f"=== PHÂN TÍCH VĂN BẢN TIẾNG VIỆT (VietAIDetector + ViDetect + viet-nhu-nguoi): {file_path} ===")
    print(f"- Số đoạn văn hợp lệ sau lọc (count_words >= 5, special_char_ratio <= 0.50): {len(norm_paras)}")
    stats = res["stats"]
    print(
        f"- Điểm dấu hiệu bề mặt (0-100): {res['surf_score']}/100 ({res['surf_band']}) "
        f"| Mẫu AI/văn dịch: {res['pattern_hit_count']} | Chữ đệm: {res['filler_per_1k']}/1k từ | Bị động: {res['passive_per_1k']}/1k từ"
    )
    if stats:
        print(
            f"- Tổng số từ văn xuôi: {stats['total_words']} | "
            f"TB: {stats['avg_sents_per_para']} câu/đoạn ({stats['avg_words_per_para']} từ/đoạn) | "
            f"Nhịp câu: {stats['avg_words_per_sent']} từ/câu (CV={stats['cv_sent_len']})"
        )
        print(
            f"- Phân bố đoạn: Đoạn 1-3 câu = {stats['short_2_3_sent_para_ratio']}%, "
            f"Đoạn 4-5 câu = {stats['uniform_4_5_sent_para_ratio']}% | Min={stats['min_words_para']} từ, Max={stats['max_words_para']} từ"
        )

    # Kiểm tra model cục bộ vietnamese-ai-detection-v2
    script_dir = os.path.dirname(os.path.abspath(__file__))
    workspace_dir = os.path.dirname(os.path.dirname(script_dir))
    local_model_dir = os.path.join(workspace_dir, "vietnamese-ai-detection-v2")
    if not os.path.exists(local_model_dir):
        local_model_dir = r"a:\agent viet bao cao\vietnamese-ai-detection-v2"

    try:
        import torch
        from transformers import AutoModelForSequenceClassification, AutoTokenizer

        if os.path.exists(os.path.join(local_model_dir, "model.safetensors")):
            print(f"- Đang chấm điểm bằng model PhoBERT cục bộ tại: {local_model_dir} ...")
            tokenizer = AutoTokenizer.from_pretrained(local_model_dir)
            model = AutoModelForSequenceClassification.from_pretrained(local_model_dir).eval()
            sample_paras = [p for p in norm_paras if count_words_vi(p) >= 20][:30]
            probs_list = []
            with torch.no_grad():
                for p in sample_paras:
                    enc = tokenizer(p, return_tensors="pt", truncation=True, max_length=256)
                    logits = model(**enc).logits
                    probs = torch.softmax(logits, dim=-1)[0].tolist()
                    probs_list.append(probs)
            if probs_list:
                avg_p0 = sum(x[0] for x in probs_list) / len(probs_list)
                avg_p1 = sum(x[1] for x in probs_list) / len(probs_list)
                print(
                    f"- Kết quả PhoBERT trên {len(probs_list)} đoạn (<=256 tokens): "
                    f"TB LABEL_0 = {avg_p0:.3f}, TB LABEL_1 = {avg_p1:.3f}"
                )
    except ImportError:
        print(
            "- [Ghi chú] Môi trường hiện chưa cài 'torch' và 'transformers' nên bỏ qua bước chạy neural PhoBERT; "
            "đã hoàn tất kiểm tra chuẩn hóa VietAIDetector & thống kê ViDetect."
        )


# ==============================================================================
# 4. CÁC HÀM GHI TRỰC TIẾP VÀO FILE .DOCX (Hnhv, Tnhnhv, Heading, Table, Code)
# ==============================================================================
REQUIRED_STYLE_IDS = {
    "Heading1", "Heading2", "Heading3", "Heading4",
    "UH1", "UH2", "Normal", "NoIndent", "Center", "Code",
    "Hnhv", "Tnhnhv", "Tnbng", "Table-Header", "Table", "Table-Center", "Table-Right",
    "Tiliuthamkho", "Footer",
}


def ensure_kma_styles_in_docx(styles_root: ET.Element, sync_existing: bool = False):
    """
    Đảm bảo file .docx đích có đầy đủ bộ style chuẩn ĐATN (Normal, Hnhv, Tnhnhv, Tnbng, Code, Footer, UH1...).
    - Nếu sync_existing=True (chế độ apply-format): đồng bộ lại font Times New Roman 14pt, căn lề,
      giãn dòng, thụt đầu dòng cho các style đã tồn tại theo đúng Mau_trinh_bay_DA.docx,
      nhưng vẫn giữ nguyên numPr sẵn có của file gốc để không phá hỏng đánh số chương mục.
    """
    existing_map = {}
    auto_numbered_styles = set()
    for st in styles_root.findall("w:style", NS):
        sid = st.attrib.get(W_ID, "")
        existing_map[sid] = st
        if st.find(".//w:numPr", NS) is not None:
            auto_numbered_styles.add(sid)

    tpl_path = get_template_path()
    if not os.path.exists(tpl_path):
        return auto_numbered_styles

    with zipfile.ZipFile(tpl_path, "r") as tz:
        tpl_styles = ET.fromstring(tz.read("word/styles.xml"))
        for tpl_st in tpl_styles.findall("w:style", NS):
            sid = tpl_st.attrib.get(W_ID, "")
            if sid not in REQUIRED_STYLE_IDS:
                continue
            if sid not in existing_map:
                st_copy = ET.fromstring(ET.tostring(tpl_st, encoding="utf-8"))
                for ppr in st_copy.findall("w:pPr", NS):
                    numpr = ppr.find("w:numPr", NS)
                    if numpr is not None:
                        ppr.remove(numpr)
                styles_root.append(st_copy)
                existing_map[sid] = st_copy
            elif sync_existing:
                cur_st = existing_map[sid]
                orig_numpr = None
                cur_ppr = cur_st.find("w:pPr", NS)
                if cur_ppr is not None:
                    np_el = cur_ppr.find("w:numPr", NS)
                    if np_el is not None:
                        orig_numpr = ET.fromstring(ET.tostring(np_el, encoding="utf-8"))

                tpl_copy = ET.fromstring(ET.tostring(tpl_st, encoding="utf-8"))
                # Giữ nguyên numPr gốc (hoặc không có numPr) của file đích và sắp xếp đúng thứ tự XSD CT_PPrBase
                for ppr in tpl_copy.findall("w:pPr", NS):
                    np_el = ppr.find("w:numPr", NS)
                    if np_el is not None:
                        ppr.remove(np_el)
                    if orig_numpr is not None:
                        ppr.append(orig_numpr)
                    _sort_ppr_children(ppr)
                if orig_numpr is not None and tpl_copy.find("w:pPr", NS) is None:
                    new_ppr = make_w_el("pPr")
                    new_ppr.append(orig_numpr)
                    tpl_copy.append(new_ppr)

                # Với Normal, khóa chặt before="0" after="0" để không bị docDefaults cũ làm giãn đoạn
                if sid == "Normal":
                    ppr_n = tpl_copy.find("w:pPr", NS)
                    if ppr_n is not None:
                        sp_n = ppr_n.find("w:spacing", NS)
                        if sp_n is not None:
                            sp_n.set(f"{{{W_URI}}}before", "0")
                            sp_n.set(f"{{{W_URI}}}after", "0")

                # Thay thế các thẻ định dạng (basedOn, pPr, rPr) trên cur_st nhưng giữ thuộc tính gốc
                for tag_name in ("basedOn", "next", "link", "pPr", "rPr"):
                    old_el = cur_st.find(f"w:{tag_name}", NS)
                    if old_el is not None:
                        cur_st.remove(old_el)
                    new_el = tpl_copy.find(f"w:{tag_name}", NS)
                    if new_el is not None:
                        cur_st.append(new_el)

    return auto_numbered_styles


def _build_kma_footer_xml(with_page_num: bool) -> bytes:
    """Tạo XML chân trang chuẩn ĐATN: trống (trang bìa) hoặc đánh số trang căn giữa Times New Roman 14pt."""
    ftr = make_w_el("ftr")
    p = make_w_el("p")
    ppr = make_w_el("pPr")
    ppr.append(make_w_el("pStyle", {"val": "Footer"}))
    if with_page_num:
        ppr.append(make_w_el("jc", {"val": "center"}))
    p.append(ppr)
    if with_page_num:
        def _ftr_rpr():
            rpr = make_w_el("rPr")
            rpr.append(make_w_el("rFonts", {"ascii": "Times New Roman", "hAnsi": "Times New Roman", "cs": "Times New Roman"}))
            rpr.append(make_w_el("sz", {"val": "28"}))
            rpr.append(make_w_el("szCs", {"val": "28"}))
            return rpr

        for ftype in ("begin",):
            r = make_w_el("r")
            r.append(_ftr_rpr())
            r.append(make_w_el("fldChar", {"fldCharType": ftype}))
            p.append(r)

        r_ins = make_w_el("r")
        r_ins.append(_ftr_rpr())
        it = make_w_el("instrText")
        it.set(XML_SPACE, "preserve")
        it.text = " PAGE   \\* MERGEFORMAT "
        r_ins.append(it)
        p.append(r_ins)

        r_sep = make_w_el("r")
        r_sep.append(_ftr_rpr())
        r_sep.append(make_w_el("fldChar", {"fldCharType": "separate"}))
        p.append(r_sep)

        r_val = make_w_el("r")
        r_val.append(_ftr_rpr())
        t_val = make_w_el("t")
        t_val.text = "1"
        r_val.append(t_val)
        p.append(r_val)

        r_end = make_w_el("r")
        r_end.append(_ftr_rpr())
        r_end.append(make_w_el("fldChar", {"fldCharType": "end"}))
        p.append(r_end)

    _sort_ppr_children(ppr)
    ftr.append(p)
    return serialize_ooxml_part(ftr)


def _set_sectpr_child_ordered(sp: ET.Element, child: ET.Element):
    """Chèn thẻ con vào w:sectPr theo đúng thứ tự chuẩn OOXML."""
    order = [
        "headerReference", "footerReference", "footnotePr", "endnotePr",
        "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
        "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
        "textDirection", "bidi", "rtlGutter", "docGrid",
    ]
    local = child.tag.split("}")[-1]
    target_rank = order.index(local) if local in order else len(order)
    for i, existing in enumerate(list(sp)):
        ex_local = existing.tag.split("}")[-1]
        ex_rank = order.index(ex_local) if ex_local in order else len(order)
        if ex_rank > target_rank:
            sp.insert(i, child)
            return
    sp.append(child)


def _apply_kma_page_and_footers(zin: zipfile.ZipFile, doc_root: ET.Element, rels_root: ET.Element, ct_root: ET.Element):
    """
    Áp dụng khổ giấy A4, căn lề chuẩn KMA (2.0 - 2.0 - 3.0 - 1.5 cm, footer 1.25 cm)
    và chuẩn hóa chân trang đánh số trang căn giữa (Times New Roman 14pt) cho mọi section.
    Trả về dict các file XML footer mới cần ghi bổ sung vào gói .docx.
    """
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    FOOTER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"

    body = doc_root.find("w:body", NS)
    children = list(body)

    # Tách các section dựa trên p/pPr/sectPr và body/sectPr cuối
    sections = []
    cur_paras = []
    for ch in children:
        tag = ch.tag.split("}")[-1]
        if tag == "p":
            cur_paras.append(ch)
            ppr = ch.find("w:pPr", NS)
            if ppr is not None:
                sp = ppr.find("w:sectPr", NS)
                if sp is not None:
                    sections.append((cur_paras, sp))
                    cur_paras = []
        elif tag == "sectPr":
            sections.append((cur_paras, ch))
            cur_paras = []
        else:
            cur_paras.extend(ch.findall(".//w:p", NS))

    if not sections:
        return {}

    # Phân loại từng section: cover / front / body
    front_title_kws = (
        "mục lục", "danh mục", "lời cảm ơn", "lời nói đầu", "lời mở đầu",
        "tóm tắt", "một số lời khuyên", "kí hiệu", "ký hiệu", "chữ viết tắt",
    )
    roles = []
    in_body = False
    for s_idx, (paras, _sp) in enumerate(sections):
        sec_texts = []
        has_body_h1 = False
        for p in paras:
            ppr = p.find("w:pPr", NS)
            st = "Normal"
            if ppr is not None:
                ps = ppr.find("w:pStyle", NS)
                if ps is not None:
                    st = ps.attrib.get(W_VAL, "Normal")
            t = "".join(x.text for x in p.findall(".//w:t", NS) if x.text).strip()
            if not t:
                continue
            low = t.lower()
            sec_texts.append(low)
            if st.startswith("TOC") or st in ("Mclc", "TableofFigures", "UH1", "UH2", "TrangBa"):
                continue
            if st == "Heading1" and not any(k in low for k in front_title_kws):
                has_body_h1 = True
            elif (
                st == "Normal"
                and not any(k in " ".join(sec_texts[:-1]) for k in front_title_kws)
                and re.match(r"^(?:chương\s+(?:1|i)[\.\:\s]|ii\.\s*nội\s+dung)", low)
                and not re.search(r"\d+$", t)
            ):
                has_body_h1 = True

        joined = " ".join(sec_texts)
        if in_body or has_body_h1:
            in_body = True
            roles.append("body")
        elif s_idx == 0 and len(sections) >= 2 and (
            any(k in joined for k in ("học viện", "báo cáo", "đồ án", "sinh viên thực hiện"))
            or not joined
        ):
            roles.append("cover")
        elif s_idx == 0 and len(sections) == 1:
            roles.append("single")
        else:
            roles.append("front")

    # Đảm bảo ít nhất section cuối được coi là body nếu chưa có body/single
    if "body" not in roles and "single" not in roles:
        roles[-1] = "body"

    # Tìm ID rId trống không trùng trong rels_root
    existing_rids = {el.attrib.get("Id", "") for el in rels_root}
    max_num = 100
    for rid in existing_rids:
        m = re.match(r"rId(\d+)$", rid)
        if m:
            max_num = max(max_num, int(m.group(1)))

    def _ensure_rel(target_name: str) -> str:
        nonlocal max_num
        for el in rels_root:
            if el.attrib.get("Target") == target_name and el.attrib.get("Type") == FOOTER_REL_TYPE:
                return el.attrib["Id"]
        max_num += 1
        new_rid = f"rId{max_num}"
        rel_el = ET.Element(f"{{{REL_NS}}}Relationship")
        rel_el.set("Id", new_rid)
        rel_el.set("Type", FOOTER_REL_TYPE)
        rel_el.set("Target", target_name)
        rels_root.append(rel_el)
        return new_rid

    def _ensure_ct(part_name: str):
        for el in ct_root:
            if el.tag.endswith("Override") and el.attrib.get("PartName") == part_name:
                return
        ov = ET.Element(f"{{{CT_NS}}}Override")
        ov.set("PartName", part_name)
        ov.set("ContentType", FOOTER_CT)
        ct_root.append(ov)

    rid_empty = _ensure_rel("footer_kma_empty.xml")
    rid_page = _ensure_rel("footer_kma_page.xml")
    _ensure_ct("/word/footer_kma_empty.xml")
    _ensure_ct("/word/footer_kma_page.xml")

    extra_parts = {
        "word/footer_kma_empty.xml": _build_kma_footer_xml(with_page_num=False),
        "word/footer_kma_page.xml": _build_kma_footer_xml(with_page_num=True),
    }

    # Map cũ từ rels để kiểm tra header trống hay có nội dung
    old_rels = {el.attrib.get("Id", ""): el.attrib.get("Target", "") for el in rels_root}

    front_started = False
    body_started = False
    for role, (_paras, sp) in zip(roles, sections):
        # 1. Khổ giấy A4 (giữ portrait nếu chưa phải landscape)
        pgsz = sp.find("w:pgSz", NS)
        if pgsz is None:
            pgsz = make_w_el("pgSz")
            _set_sectpr_child_ordered(sp, pgsz)
        if pgsz.attrib.get(f"{{{W_URI}}}orient") != "landscape":
            pgsz.set(f"{{{W_URI}}}w", KMA_PAGE_W)
            pgsz.set(f"{{{W_URI}}}h", KMA_PAGE_H)

        # 2. Căn lề chuẩn ĐATN
        pgmar = sp.find("w:pgMar", NS)
        if pgmar is None:
            pgmar = make_w_el("pgMar")
            _set_sectpr_child_ordered(sp, pgmar)
        for mk, mv in KMA_MARGINS.items():
            pgmar.set(f"{{{W_URI}}}{mk}", mv)

        # 3. Xóa footerReference cũ và headerReference trống (nếu có)
        for fr in list(sp.findall("w:footerReference", NS)):
            sp.remove(fr)
        for hr in list(sp.findall("w:headerReference", NS)):
            htarget = old_rels.get(hr.attrib.get(f"{{{R_NS}}}id", ""), "")
            hpart = f"word/{htarget}" if htarget and not htarget.startswith("word/") else htarget
            if hpart and hpart in zin.namelist():
                hroot = ET.fromstring(zin.read(hpart))
                htxt = "".join(t.text for t in hroot.findall(".//w:t", NS) if t.text).strip()
                if not htxt and hroot.find(".//w:drawing", NS) is None:
                    sp.remove(hr)

        for old_tag in ("w:pgNumType", "w:titlePg"):
            old_el = sp.find(old_tag, NS)
            if old_el is not None:
                sp.remove(old_el)

        # 4. Gắn chân trang & số trang đúng vai trò section
        if role == "cover":
            fr = make_w_el("footerReference", {"type": "default"})
            fr.set(f"{{{R_NS}}}id", rid_empty)
            _set_sectpr_child_ordered(sp, fr)
        elif role == "front":
            fr = make_w_el("footerReference", {"type": "default"})
            fr.set(f"{{{R_NS}}}id", rid_page)
            _set_sectpr_child_ordered(sp, fr)
            pgnum = make_w_el("pgNumType", {"fmt": "lowerRoman"})
            if not front_started:
                pgnum.set(f"{{{W_URI}}}start", "1")
                front_started = True
            _set_sectpr_child_ordered(sp, pgnum)
        elif role == "body":
            fr = make_w_el("footerReference", {"type": "default"})
            fr.set(f"{{{R_NS}}}id", rid_page)
            _set_sectpr_child_ordered(sp, fr)
            if not body_started:
                pgnum = make_w_el("pgNumType", {"start": "1"})
                _set_sectpr_child_ordered(sp, pgnum)
                body_started = True
        else:  # single-section doc
            fr_def = make_w_el("footerReference", {"type": "default"})
            fr_def.set(f"{{{R_NS}}}id", rid_page)
            _set_sectpr_child_ordered(sp, fr_def)
            fr_fst = make_w_el("footerReference", {"type": "first"})
            fr_fst.set(f"{{{R_NS}}}id", rid_empty)
            _set_sectpr_child_ordered(sp, fr_fst)
            _set_sectpr_child_ordered(sp, make_w_el("titlePg"))
            _set_sectpr_child_ordered(sp, make_w_el("pgNumType", {"start": "1"}))

    return extra_parts


def _normalize_body_paragraphs_format(body: ET.Element):
    """
    Chuẩn hóa font chữ (Times New Roman 14pt cho nội dung, Courier New 12pt cho Code)
    và căn đều / giãn dòng / thụt đầu dòng cho các đoạn văn thân bài (bỏ qua Trang bìa và Mục lục).
    """
    special_keep_fonts = {"Symbol", "Wingdings", "Webdings", "Cambria Math"}
    body_children = list(body)

    # Xác định ranh giới hết trang bìa (hết section đầu tiên nếu có >= 2 section, hoặc tới trước Mục lục/Lời mở đầu/Chương 1)
    sect_breaks = [i for i, ch in enumerate(body_children) if ch.tag == f"{{{W_URI}}}p" and ch.find(".//w:sectPr", NS) is not None]
    cover_end_idx = sect_breaks[0] if sect_breaks else -1

    in_cover = cover_end_idx >= 0
    for idx, ch in enumerate(body_children):
        if cover_end_idx >= 0 and idx > cover_end_idx:
            in_cover = False

        for p in ([ch] if ch.tag == f"{{{W_URI}}}p" else ch.findall(".//w:p", NS)):
            ppr = p.find("w:pPr", NS)
            st = "Normal"
            if ppr is not None:
                ps = ppr.find("w:pStyle", NS)
                if ps is not None:
                    st = ps.attrib.get(W_VAL, "Normal")

            txt = "".join(t.text for t in p.findall(".//w:t", NS) if t.text).strip()
            low_txt = txt.lower()
            if st in ("UH1", "Heading1", "Mclc") or any(k in low_txt for k in ("mục lục", "lời nói đầu", "lời mở đầu", "tóm tắt nội dung", "chương 1", "chương i")):
                in_cover = False

            if in_cover or st == "TrangBa" or st.startswith("TOC") or st in ("Mclc", "TableofFigures"):
                continue

            # Chuẩn hóa runs trong đoạn văn thân bài
            for r in p.findall("w:r", NS):
                rpr = r.find("w:rPr", NS)
                if rpr is None:
                    continue
                if rpr.find("w:vertAlign", NS) is not None:
                    continue
                rf = rpr.find("w:rFonts", NS)
                if rf is not None:
                    cur_f = rf.attrib.get(f"{{{W_URI}}}ascii") or rf.attrib.get(f"{{{W_URI}}}hAnsi") or ""
                    if cur_f in special_keep_fonts:
                        continue
                    target_font = "Courier New" if st == "Code" else "Times New Roman"
                    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                        rf.set(f"{{{W_URI}}}{attr}", target_font)
                    for theme_attr in ("asciiTheme", "hAnsiTheme", "csTheme", "eastAsiaTheme"):
                        full_ta = f"{{{W_URI}}}{theme_attr}"
                        if full_ta in rf.attrib:
                            del rf.attrib[full_ta]

                sz = rpr.find("w:sz", NS)
                szcs = rpr.find("w:szCs", NS)
                if st == "Code":
                    cur_sz = sz.attrib.get(f"{{{W_URI}}}val", "") if sz is not None else ""
                    target_sz = cur_sz if cur_sz in ("20", "21", "22", "24") else "21"
                    if sz is not None:
                        sz.set(f"{{{W_URI}}}val", target_sz)
                    if szcs is not None:
                        szcs.set(f"{{{W_URI}}}val", target_sz)
                elif st in ("Normal", "NoIndent", "Center", "Tiliuthamkho", "Tnhnhv", "Tnbng", "Table", "Table-Header", "Table-Center", "Table-Right", "Heading1", "Heading2", "Heading3", "Heading4", "UH1", "UH2"):
                    if sz is not None:
                        rpr.remove(sz)
                    if szcs is not None:
                        rpr.remove(szcs)

            # Gỡ spacing/indent/jc lệch trên đoạn văn xuôi Normal (giữ nguyên bên trong bảng, bullet, khung giữ chỗ ảnh, căn giữa/phải)
            if ch.tag == f"{{{W_URI}}}p" and st == "Normal" and ppr is not None:
                is_box = ppr.find("w:pBdr", NS) is not None or ppr.find("w:shd", NS) is not None
                has_num = ppr.find("w:numPr", NS) is not None
                jc_el = ppr.find("w:jc", NS)
                jc_val = jc_el.attrib.get(W_VAL) if jc_el is not None else None
                is_bullet_or_cmd = txt.startswith(("-", "+", "•", "❖", "$ ", "# ", "Bước ")) or PLACEHOLDER_MARKER_RE.search(txt)
                has_drawing = p.find(".//w:drawing", NS) is not None or p.find(".//w:pict", NS) is not None
                if not is_box and not has_num and not is_bullet_or_cmd and not has_drawing and jc_val not in ("center", "right"):
                    if jc_el is not None:
                        jc_el.set(W_VAL, "both")
                    sp_el = ppr.find("w:spacing", NS)
                    if sp_el is not None:
                        ppr.remove(sp_el)
                    ind_el = ppr.find("w:ind", NS)
                    if ind_el is not None and count_words_vi(txt) >= 12:
                        ppr.remove(ind_el)


def _normalize_and_align_images(
    zin: zipfile.ZipFile,
    body: ET.Element,
    rels_root: ET.Element,
    max_width_cm: float = 15.5,
    max_height_cm: float = 11.0,
) -> int:
    """
    Tự động căn chỉnh toàn bộ ảnh đã chèn trong thân bài (bỏ qua Trang bìa):
    - Đưa đoạn chứa ảnh về đúng style 'Hnhv' (thay vì bị dính nhầm style caption 'Tnhnhv' hoặc 'Normal' thụt lề 1.25cm).
    - Gỡ khung nét đứt/nền xám giữ chỗ cũ nếu người dùng dán ảnh trực tiếp vào khung giữ chỗ.
    - Khóa căn giữa (jc=center), firstLine=0, giãn dòng đơn (line=240) và keepNext để ảnh luôn đi liền với dòng Tên hình bên dưới.
    - Co giãn tỷ lệ khung hình chuẩn xác theo kích thước điểm ảnh thực tế (rộng <= max_width_cm, cao <= max_height_cm)
      đồng bộ trên cả <wp:extent> lẫn <pic:spPr><a:xfrm><a:ext>, đồng thời khử trùng lặp id của wp:docPr.
    """
    rels_map = {el.attrib.get("Id", ""): el.attrib.get("Target", "") for el in rels_root}
    media_dims = {}
    names_in_zip = set(zin.namelist())
    for rid, target in rels_map.items():
        if not target:
            continue
        mpath = f"word/{target.lstrip('/')}" if not target.startswith("word/") else target
        if mpath in names_in_zip and mpath.startswith("word/media/"):
            dims = get_image_pixel_dimensions(zin.read(mpath))
            if dims:
                media_dims[rid] = dims

    # Bước 0: Nếu người dùng dán ảnh trực tiếp vào đầu đoạn chú thích Tên hình (cùng 1 thẻ <w:p> vừa có w:drawing vừa có 'Hình X.Y...'),
    # tự động tách thành 2 đoạn riêng: đoạn trên chứa ảnh (style Hnhv) và đoạn dưới giữ nguyên trường động Tên hình (style Tnhnhv).
    raw_children = list(body)
    raw_sect_breaks = [
        i for i, ch in enumerate(raw_children)
        if ch.tag == f"{{{W_URI}}}p" and ch.find(".//w:sectPr", NS) is not None
    ]
    raw_cover_end = raw_sect_breaks[0] if raw_sect_breaks else -1
    in_cov_split = raw_cover_end >= 0
    for ch in raw_children:
        idx_now = list(body).index(ch)
        if raw_cover_end >= 0 and idx_now > raw_cover_end:
            in_cov_split = False
        if ch.tag != f"{{{W_URI}}}p":
            continue
        ppr0 = ch.find("w:pPr", NS)
        st0 = "Normal"
        if ppr0 is not None:
            ps0 = ppr0.find("w:pStyle", NS)
            if ps0 is not None:
                st0 = ps0.attrib.get(W_VAL, "Normal")
        txt0 = "".join(t.text for t in ch.findall(".//w:t", NS) if t.text).strip()
        if st0 in ("UH1", "Heading1", "Mclc") or any(
            k in txt0.lower() for k in ("mục lục", "lời nói đầu", "lời mở đầu", "tóm tắt nội dung", "chương 1", "chương i")
        ):
            in_cov_split = False
        if in_cov_split or st0 == "TrangBa":
            continue
        has_media0 = ch.find(".//w:drawing", NS) is not None or ch.find(".//w:pict", NS) is not None
        if has_media0 and not PLACEHOLDER_MARKER_RE.search(txt0) and re.match(r"(?i)^hình\s*\d+", txt0):
            img_runs = [
                sub for sub in list(ch)
                if sub.tag == f"{{{W_URI}}}r"
                and (sub.find(".//w:drawing", NS) is not None or sub.find(".//w:pict", NS) is not None)
                and not "".join(t.text for t in sub.findall(".//w:t", NS) if t.text).strip()
            ]
            if img_runs:
                p_img = make_w_el("p")
                ppr_img = make_w_el("pPr")
                ppr_img.append(make_w_el("pStyle", {"val": "Hnhv"}))
                p_img.append(ppr_img)
                for ir in img_runs:
                    ch.remove(ir)
                    p_img.append(ir)
                if ppr0 is None:
                    ppr0 = make_w_el("pPr")
                    ch.insert(0, ppr0)
                ps0 = ppr0.find("w:pStyle", NS)
                if ps0 is None:
                    ppr0.append(make_w_el("pStyle", {"val": "Tnhnhv"}))
                else:
                    ps0.set(W_VAL, "Tnhnhv")
                _sort_ppr_children(ppr0)
                pos_in_body = list(body).index(ch)
                body.insert(pos_in_body, p_img)

    body_children = list(body)
    sect_breaks = [
        i for i, ch in enumerate(body_children)
        if ch.tag == f"{{{W_URI}}}p" and ch.find(".//w:sectPr", NS) is not None
    ]
    cover_end_idx = sect_breaks[0] if sect_breaks else -1

    # Cấp phát lại ID duy nhất cho toàn bộ đối tượng wp:docPr trong tài liệu để tránh lỗi trùng ID của Word
    next_docpr_id = 1
    for dp in body.findall(f".//{{{WP_NS}}}docPr"):
        dp.set("id", str(next_docpr_id))
        next_docpr_id += 1

    EMU_PER_CM = 360000
    max_w_emu = int(round(max_width_cm * EMU_PER_CM))
    max_h_emu = int(round(max_height_cm * EMU_PER_CM))
    aligned_count = 0

    in_cover = cover_end_idx >= 0
    for idx, ch in enumerate(body_children):
        if cover_end_idx >= 0 and idx > cover_end_idx:
            in_cover = False
        if ch.tag != f"{{{W_URI}}}p":
            continue

        ppr = ch.find("w:pPr", NS)
        st = "Normal"
        if ppr is not None:
            ps = ppr.find("w:pStyle", NS)
            if ps is not None:
                st = ps.attrib.get(W_VAL, "Normal")

        txt = "".join(t.text for t in ch.findall(".//w:t", NS) if t.text).strip()
        low_txt = txt.lower()
        if st in ("UH1", "Heading1", "Mclc") or any(
            k in low_txt for k in ("mục lục", "lời nói đầu", "lời mở đầu", "tóm tắt nội dung", "chương 1", "chương i")
        ):
            in_cover = False

        if in_cover or st == "TrangBa":
            continue

        drawings = ch.findall(".//w:drawing", NS)
        picts = ch.findall(".//w:pict", NS)
        if not drawings and not picts:
            continue

        # Chuẩn hóa đoạn văn chứa ảnh về đúng style Hnhv chuẩn KMA
        if ppr is None:
            ppr = make_w_el("pPr")
            ch.insert(0, ppr)

        ps = ppr.find("w:pStyle", NS)
        if ps is None:
            ps = make_w_el("pStyle", {"val": "Hnhv"})
            ppr.append(ps)
        else:
            ps.set(W_VAL, "Hnhv")

        # Xóa khung viền/nền giữ chỗ cũ hoặc thuộc tính đánh số nếu còn sót
        for old_tag in ("w:pBdr", "w:shd", "w:numPr"):
            old_el = ppr.find(old_tag, NS)
            if old_el is not None:
                ppr.remove(old_el)

        # Gỡ chữ hướng dẫn [KHUNG HÌNH ...] nếu người dùng dán ảnh đè chung vào đoạn giữ chỗ
        if PLACEHOLDER_MARKER_RE.search(txt):
            for r in list(ch.findall("w:r", NS)):
                if r.find(".//w:drawing", NS) is None and r.find(".//w:pict", NS) is None:
                    ch.remove(r)

        if ppr.find("w:keepNext", NS) is None:
            ppr.append(make_w_el("keepNext"))

        sp_el = ppr.find("w:spacing", NS)
        if sp_el is None:
            sp_el = make_w_el("spacing")
            ppr.append(sp_el)
        sp_el.attrib.clear()
        sp_el.set(f"{{{W_URI}}}before", "120")
        sp_el.set(f"{{{W_URI}}}after", "60")
        sp_el.set(f"{{{W_URI}}}line", "240")
        sp_el.set(f"{{{W_URI}}}lineRule", "auto")

        ind_el = ppr.find("w:ind", NS)
        if ind_el is None:
            ind_el = make_w_el("ind")
            ppr.append(ind_el)
        ind_el.attrib.clear()
        ind_el.set(f"{{{W_URI}}}left", "0")
        ind_el.set(f"{{{W_URI}}}right", "0")
        ind_el.set(f"{{{W_URI}}}firstLine", "0")

        jc_el = ppr.find("w:jc", NS)
        if jc_el is None:
            jc_el = make_w_el("jc", {"val": "center"})
            ppr.append(jc_el)
        else:
            jc_el.set(W_VAL, "center")

        # Gỡ in nghiêng/đậm thừa trên run bọc ảnh (do dính từ Tnhnhv)
        for r in ch.findall("w:r", NS):
            if r.find(".//w:drawing", NS) is not None or r.find(".//w:pict", NS) is not None:
                rpr = r.find("w:rPr", NS)
                if rpr is not None:
                    for fmt_tag in ("w:i", "w:iCs", "w:b", "w:bCs"):
                        fe = rpr.find(fmt_tag, NS)
                        if fe is not None:
                            rpr.remove(fe)

        _sort_ppr_children(ppr)

        # Căn chỉnh kích thước từng đối tượng w:drawing giữ nguyên tỷ lệ khung hình gốc
        for d in drawings:
            ext_el = d.find(f".//{{{WP_NS}}}extent")
            if ext_el is None:
                continue
            cur_cx = int(ext_el.attrib.get("cx", "0") or "0")
            cur_cy = int(ext_el.attrib.get("cy", "0") or "0")

            blip = d.find(f".//{{{A_NS}}}blip")
            rid = blip.attrib.get(f"{{{R_NS_DOC}}}embed", "") if blip is not None else ""
            px_dims = media_dims.get(rid)

            if px_dims:
                w_px, h_px = px_dims
                ratio = w_px / float(h_px)
                # Kích thước vật lý gốc ở 96 DPI (1 px = 9525 EMU); không phóng đại ảnh icon nhỏ quá 2 lần
                nat_cx = w_px * 9525
                target_cx = min(max_w_emu, max(nat_cx, int(max_w_emu * 0.65)) if w_px >= 450 else min(max_w_emu, nat_cx * 2))
                target_cy = int(round(target_cx / ratio))
                if target_cy > max_h_emu:
                    target_cy = max_h_emu
                    target_cx = int(round(target_cy * ratio))
            elif cur_cx > 0 and cur_cy > 0:
                ratio = cur_cx / float(cur_cy)
                target_cx = min(cur_cx, max_w_emu)
                target_cy = int(round(target_cx / ratio))
                if target_cy > max_h_emu:
                    target_cy = max_h_emu
                    target_cx = int(round(target_cy * ratio))
            else:
                continue

            ext_el.set("cx", str(target_cx))
            ext_el.set("cy", str(target_cy))
            for a_ext in d.findall(f".//{{{PIC_NS}}}spPr/{{{A_NS}}}xfrm/{{{A_NS}}}ext"):
                a_ext.set("cx", str(target_cx))
                a_ext.set("cy", str(target_cy))
            aligned_count += 1

        # Đảm bảo dòng chú thích Hình ngay bên dưới dùng đúng style Tnhnhv căn giữa
        if idx + 1 < len(body_children) and body_children[idx + 1].tag == f"{{{W_URI}}}p":
            nxt_p = body_children[idx + 1]
            nxt_txt = "".join(t.text for t in nxt_p.findall(".//w:t", NS) if t.text).strip()
            if re.match(r"(?i)^hình\s+\d+", nxt_txt):
                nxt_ppr = nxt_p.find("w:pPr", NS)
                if nxt_ppr is None:
                    nxt_ppr = make_w_el("pPr")
                    nxt_p.insert(0, nxt_ppr)
                nxt_ps = nxt_ppr.find("w:pStyle", NS)
                if nxt_ps is None:
                    nxt_ppr.append(make_w_el("pStyle", {"val": "Tnhnhv"}))
                else:
                    nxt_ps.set(W_VAL, "Tnhnhv")
                _sort_ppr_children(nxt_ppr)

    return aligned_count


def cmd_apply_format(docx_path: str, max_img_width_cm: float = 15.5, max_img_height_cm: float = 11.0):
    """
    Áp dụng đồng bộ toàn bộ Format chuẩn ĐATN (Mau_trinh_bay_DA.docx) cho file .docx
    cả ở chế độ Bài tập lớn (btl) và Đồ án tốt nghiệp (do-an):
    khổ A4, lề 2.0-2.0-3.0-1.5cm, chân trang đánh số căn giữa 1.25cm, bộ style & font Times New Roman 14pt,
    tự động căn chỉnh hình ảnh về style Hnhv căn giữa và xuất gói XML chuẩn Microsoft Word không lỗi khôi phục.
    """
    if not os.path.exists(docx_path):
        print(f"[LỖI] Không tìm thấy file: {docx_path}")
        sys.exit(1)

    REL_DEFAULT_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    CT_DEFAULT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

    tmp_docx = docx_path + ".tmp"
    img_aligned = 0
    with zipfile.ZipFile(docx_path, "r") as zin:
        doc_root, doc_ns = parse_ooxml_part(zin.read("word/document.xml"))
        styles_root, styles_ns = parse_ooxml_part(zin.read("word/styles.xml"))
        rels_root, rels_ns = parse_ooxml_part(zin.read("word/_rels/document.xml.rels"))
        ct_root, ct_ns = parse_ooxml_part(zin.read("[Content_Types].xml"))

        ensure_kma_styles_in_docx(styles_root, sync_existing=True)
        extra_parts = _apply_kma_page_and_footers(zin, doc_root, rels_root, ct_root)
        body = doc_root.find("w:body", NS)
        if body is not None:
            _normalize_body_paragraphs_format(body)
            wrap_body_code_blocks_in_frames(body)
            img_aligned = _normalize_and_align_images(
                zin, body, rels_root, max_width_cm=max_img_width_cm, max_height_cm=max_img_height_cm
            )

        _sort_all_pprs(styles_root)
        _sort_all_pprs(doc_root)

        doc_bytes = serialize_ooxml_part(doc_root, ns_map=doc_ns)
        styles_bytes = serialize_ooxml_part(styles_root, ns_map=styles_ns)
        rels_bytes = serialize_ooxml_part(rels_root, ns_map=rels_ns, default_ns=REL_DEFAULT_NS)
        ct_bytes = serialize_ooxml_part(ct_root, ns_map=ct_ns, default_ns=CT_DEFAULT_NS)

        with zipfile.ZipFile(tmp_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    write_clean_zip_entry(zout, item.filename, doc_bytes, item.date_time)
                elif item.filename == "word/styles.xml":
                    write_clean_zip_entry(zout, item.filename, styles_bytes, item.date_time)
                elif item.filename == "word/_rels/document.xml.rels":
                    write_clean_zip_entry(zout, item.filename, rels_bytes, item.date_time)
                elif item.filename == "[Content_Types].xml":
                    write_clean_zip_entry(zout, item.filename, ct_bytes, item.date_time)
                elif item.filename in extra_parts:
                    write_clean_zip_entry(zout, item.filename, extra_parts[item.filename], item.date_time)
                else:
                    write_clean_zip_entry(zout, item.filename, zin.read(item.filename), item.date_time)
                written.add(item.filename)
            for part_name, data in extra_parts.items():
                if part_name not in written:
                    write_clean_zip_entry(zout, part_name, data)

    ensure_docx_writable_and_replace(tmp_docx, docx_path)
    xml_issues = validate_docx_ooxml(docx_path)
    if xml_issues:
        print("[CẢNH BÁO XML]", "; ".join(xml_issues))
    print(
        f"[OK] Đã chuẩn hóa Format ĐATN KMA (A4, lề 2.0-2.0-3.0-1.5cm, chân trang căn giữa 1.25cm, "
        f"Times New Roman 14pt, căn chỉnh {img_aligned} ảnh chuẩn Hnhv, kiểm định OOXML sạch) cho: {docx_path}"
    )


def cmd_align_images(docx_path: str, max_width_cm: float = 15.5, max_height_cm: float = 11.0):
    """Chỉ căn chỉnh lại các hình ảnh đã chèn trong thân bài (căn giữa Hnhv, co giãn tỷ lệ vừa trang, giữ caption)."""
    if not os.path.exists(docx_path):
        print(f"[LỖI] Không tìm thấy file: {docx_path}")
        sys.exit(1)

    tmp_docx = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        doc_root, doc_ns = parse_ooxml_part(zin.read("word/document.xml"))
        styles_root, styles_ns = parse_ooxml_part(zin.read("word/styles.xml"))
        rels_root, _ = parse_ooxml_part(zin.read("word/_rels/document.xml.rels"))

        ensure_kma_styles_in_docx(styles_root, sync_existing=False)
        body = doc_root.find("w:body", NS)
        aligned_count = 0
        if body is not None:
            aligned_count = _normalize_and_align_images(
                zin, body, rels_root, max_width_cm=max_width_cm, max_height_cm=max_height_cm
            )
        _sort_all_pprs(styles_root)
        _sort_all_pprs(doc_root)

        doc_bytes = serialize_ooxml_part(doc_root, ns_map=doc_ns)
        styles_bytes = serialize_ooxml_part(styles_root, ns_map=styles_ns)

        with zipfile.ZipFile(tmp_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    write_clean_zip_entry(zout, item.filename, doc_bytes, item.date_time)
                elif item.filename == "word/styles.xml":
                    write_clean_zip_entry(zout, item.filename, styles_bytes, item.date_time)
                else:
                    write_clean_zip_entry(zout, item.filename, zin.read(item.filename), item.date_time)

    ensure_docx_writable_and_replace(tmp_docx, docx_path)
    print(
        f"[OK] Đã căn chỉnh {aligned_count} hình ảnh (style Hnhv căn giữa, rộng <= {max_width_cm}cm, cao <= {max_height_cm}cm) "
        f"và bảo toàn cấu trúc XML Word cho: {docx_path}"
    )


def add_run_text(p_el: ET.Element, text: str, bold: bool = False, italic: bool = False, color: str = None):
    r_el = make_w_el("r")
    if bold or italic or color:
        rpr = make_w_el("rPr")
        if bold:
            rpr.append(make_w_el("b"))
        if italic:
            rpr.append(make_w_el("i"))
        if color:
            rpr.append(make_w_el("color", {"val": color}))
        r_el.append(rpr)
    t_el = make_w_el("t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t_el.set(XML_SPACE, "preserve")
    t_el.text = text
    r_el.append(t_el)
    p_el.append(r_el)
    return r_el


def add_word_field(p_el: ET.Element, instr: str, cached_val: str):
    """Tạo trường động của Word (STYLEREF / SEQ) giống hệt Mau_trinh_bay_DA.docx."""
    r_begin = make_w_el("r")
    r_begin.append(make_w_el("fldChar", {"fldCharType": "begin"}))
    p_el.append(r_begin)

    r_instr = make_w_el("r")
    it = make_w_el("instrText")
    it.set(XML_SPACE, "preserve")
    it.text = f" {instr} "
    r_instr.append(it)
    p_el.append(r_instr)

    r_sep = make_w_el("r")
    r_sep.append(make_w_el("fldChar", {"fldCharType": "separate"}))
    p_el.append(r_sep)

    r_val = make_w_el("r")
    rpr = make_w_el("rPr")
    rpr.append(make_w_el("noProof"))
    r_val.append(rpr)
    t_val = make_w_el("t")
    t_val.text = str(cached_val)
    r_val.append(t_val)
    p_el.append(r_val)

    r_end = make_w_el("r")
    r_end.append(make_w_el("fldChar", {"fldCharType": "end"}))
    p_el.append(r_end)


def build_paragraph_xml(text: str, style_id: str = "Normal") -> ET.Element:
    p_el = make_w_el("p")
    ppr = make_w_el("pPr")
    ppr.append(make_w_el("pStyle", {"val": style_id}))
    p_el.append(ppr)
    add_run_text(p_el, text)
    return p_el


def clean_inline_markdown(text: str) -> str:
    """Gỡ sạch ký hiệu Markdown nội dòng (backtick `, **in đậm**, *in nghiêng*) để không sót vào văn xuôi/bảng Word."""
    s = re.sub(r"`([^`]+)`", r"\1", text)
    s = re.sub(r"\*\*([^*]+)\*\*", r"\1", s)
    s = re.sub(r"(?<!\w)\*([^*\n]+)\*(?!\w)", r"\1", s)
    return s


def build_image_placeholder_xml(fig_num: str, title: str, img_type: str, hint: str):
    """
    Tạo cặp 2 đoạn văn chuẩn KMA để giữ chỗ hình ảnh trong .docx:
    1. Đoạn khung giữ chỗ có style 'Hnhv', đóng khung viền nét đứt (<w:pBdr>), nền xám nhạt,
       chỉ ghi nhãn gọn [KHUNG HÌNH X.Y - CHÈN ẢNH TẠI ĐÂY] và 1 dòng mô tả kỹ thuật ngắn
       (không chèn chữ 'Gợi ý nội dung', không in tên file cục bộ .pdf/.md).
    2. Đoạn tên hình đặt ngay bên dưới có style 'Tnhnhv' kèm trường động STYLEREF & SEQ Hình.
    """
    parts = fig_num.strip().split(".", 1)
    chap_str = parts[0] if parts and parts[0] else "1"
    seq_str = parts[1] if len(parts) > 1 and parts[1] else "1"

    # Đoạn 1: Style Hnhv + Khung viền giữ chỗ
    p_box = make_w_el("p")
    ppr = make_w_el("pPr")
    ppr.append(make_w_el("pStyle", {"val": "Hnhv"}))
    pbdr = make_w_el("pBdr")
    for side in ("top", "left", "bottom", "right"):
        pbdr.append(
            make_w_el(
                side,
                {"val": "dashed", "sz": "8", "space": "6", "color": "595959"},
            )
        )
    ppr.append(pbdr)
    ppr.append(make_w_el("shd", {"val": "clear", "color": "auto", "fill": "F5F5F5"}))
    ppr.append(make_w_el("jc", {"val": "center"}))
    p_box.append(ppr)

    add_run_text(
        p_box,
        f"[KHUNG HÌNH {fig_num} - CHÈN ẢNH TẠI ĐÂY]",
        bold=True,
        color="1F4E79",
    )
    clean_hint = clean_inline_markdown(hint or "").strip()
    clean_hint = re.sub(r"\b[\w\-]+\.(?:pdf|md|html|docx)\b", "", clean_hint, flags=re.IGNORECASE).strip(" ,;:-")
    if clean_hint:
        r_br = make_w_el("r")
        r_br.append(make_w_el("br"))
        p_box.append(r_br)
        add_run_text(p_box, clean_hint, italic=True)

    # Đoạn 2: Style Tnhnhv bên dưới hình (kèm trường STYLEREF 1 \s và SEQ Hình)
    p_cap = make_w_el("p")
    ppr_cap = make_w_el("pPr")
    ppr_cap.append(make_w_el("pStyle", {"val": "Tnhnhv"}))
    p_cap.append(ppr_cap)

    add_run_text(p_cap, "Hình ")
    add_word_field(p_cap, r"STYLEREF 1 \s", chap_str)
    add_run_text(p_cap, ".")
    add_word_field(p_cap, r"SEQ Hình \* ARABIC \s 1", seq_str)
    clean_title = re.sub(r"(?i)^hình\s+\d+[\.\d]*\s*[\.:]?\s*", "", clean_inline_markdown(title).strip())
    add_run_text(p_cap, f". {clean_title}")

    return [p_box, p_cap]


def build_table_caption_xml(tbl_num: str, title: str) -> ET.Element:
    """Tạo dòng Tên bảng đặt bên trên bảng với style 'Tnbng' và trường STYLEREF / SEQ Bảng."""
    parts = tbl_num.strip().split(".", 1)
    chap_str = parts[0] if parts and parts[0] else "1"
    seq_str = parts[1] if len(parts) > 1 and parts[1] else "1"

    p_cap = make_w_el("p")
    ppr = make_w_el("pPr")
    ppr.append(make_w_el("pStyle", {"val": "Tnbng"}))
    p_cap.append(ppr)

    add_run_text(p_cap, "Bảng ")
    add_word_field(p_cap, r"STYLEREF 1 \s", chap_str)
    add_run_text(p_cap, ".")
    add_word_field(p_cap, r"SEQ Bảng \* ARABIC \s 1", seq_str)
    clean_title = re.sub(r"(?i)^bảng\s+\d+[\.\d]*\s*[\.:]?\s*", "", clean_inline_markdown(title).strip())
    add_run_text(p_cap, f". {clean_title}")
    return p_cap


def build_framed_code_table_xml(code_lines: list, font_half_pt: str = "21") -> ET.Element:
    """
    Tạo khung mã nguồn 1x1 bảng chuẩn báo cáo/ĐATN (viền trái dày xanh đậm #1E3A5F,
    viền trên/phải/dưới mảnh #D0D7DE, nền xám xanh nhạt #F6F8FA, font Courier New, tắt kiểm tra chính tả).
    Mỗi dòng bên trong giữ nguyên pStyle='Code' để tương thích 100% với bộ kiểm tra tự động.
    """
    cleaned = list(code_lines)
    while cleaned and not cleaned[0].strip():
        cleaned.pop(0)
    while cleaned and not cleaned[-1].strip():
        cleaned.pop()
    if not cleaned:
        cleaned = [""]

    tbl = make_w_el("tbl")
    tbl_pr = make_w_el("tblPr")
    tbl_pr.append(make_w_el("tblW", {"w": "9355", "type": "dxa"}))
    tbl_pr.append(make_w_el("jc", {"val": "left"}))
    tbl_pr.append(make_w_el("tblInd", {"w": "0", "type": "dxa"}))

    borders = make_w_el("tblBorders")
    borders.append(make_w_el("top", {"val": "single", "sz": "4", "space": "0", "color": "D0D7DE"}))
    borders.append(make_w_el("left", {"val": "single", "sz": "24", "space": "0", "color": "1E3A5F"}))
    borders.append(make_w_el("bottom", {"val": "single", "sz": "4", "space": "0", "color": "D0D7DE"}))
    borders.append(make_w_el("right", {"val": "single", "sz": "4", "space": "0", "color": "D0D7DE"}))
    borders.append(make_w_el("insideH", {"val": "none"}))
    borders.append(make_w_el("insideV", {"val": "none"}))
    tbl_pr.append(borders)

    cell_mar = make_w_el("tblCellMar")
    cell_mar.append(make_w_el("top", {"w": "120", "type": "dxa"}))
    cell_mar.append(make_w_el("left", {"w": "180", "type": "dxa"}))
    cell_mar.append(make_w_el("bottom", {"w": "120", "type": "dxa"}))
    cell_mar.append(make_w_el("right", {"w": "140", "type": "dxa"}))
    tbl_pr.append(cell_mar)
    tbl_pr.append(
        make_w_el(
            "tblLook",
            {
                "val": "04A0",
                "firstRow": "0",
                "lastRow": "0",
                "firstColumn": "0",
                "lastColumn": "0",
                "noHBand": "1",
                "noVBand": "1",
            },
        )
    )
    tbl.append(tbl_pr)

    grid = make_w_el("tblGrid")
    grid.append(make_w_el("gridCol", {"w": "9355"}))
    tbl.append(grid)

    tr = make_w_el("tr")
    tc = make_w_el("tc")
    tc_pr = make_w_el("tcPr")
    tc_pr.append(make_w_el("tcW", {"w": "9355", "type": "dxa"}))
    tc_pr.append(make_w_el("shd", {"val": "clear", "color": "auto", "fill": "F6F8FA"}))
    tc.append(tc_pr)

    for raw_ln in cleaned:
        ln = raw_ln.rstrip("\r\n").replace("\t", "    ")
        p = make_w_el("p")
        ppr = make_w_el("pPr")
        ppr.append(make_w_el("pStyle", {"val": "Code"}))
        ppr.append(make_w_el("spacing", {"before": "0", "after": "0", "line": "252", "lineRule": "auto"}))
        ppr.append(make_w_el("ind", {"left": "0", "right": "0", "firstLine": "0"}))
        ppr.append(make_w_el("jc", {"val": "left"}))
        _sort_ppr_children(ppr)
        p.append(ppr)

        r = make_w_el("r")
        rpr = make_w_el("rPr")
        rpr.append(
            make_w_el(
                "rFonts",
                {
                    "ascii": "Courier New",
                    "hAnsi": "Courier New",
                    "cs": "Courier New",
                    "eastAsia": "Courier New",
                },
            )
        )
        rpr.append(make_w_el("noProof"))
        rpr.append(make_w_el("color", {"val": "1F2328"}))
        rpr.append(make_w_el("sz", {"val": font_half_pt}))
        rpr.append(make_w_el("szCs", {"val": font_half_pt}))
        r.append(rpr)
        if ln:
            t = make_w_el("t")
            if ln.startswith(" ") or ln.endswith(" ") or "  " in ln:
                t.set(XML_SPACE, "preserve")
            t.text = ln
            r.append(t)
        p.append(r)
        tc.append(p)

    tr.append(tc)
    tbl.append(tr)
    return tbl


def wrap_body_code_blocks_in_frames(body: ET.Element, font_half_pt: str = "21") -> int:
    """
    Gom các đoạn <w:p style='Code'> liên tiếp nằm trực tiếp trong w:body thành từng bảng khung mã nguồn 1x1.
    An toàn lũy đẳng (không bọc chồng lên bảng đã đóng khung).
    """
    children = list(body)
    wrapped_blocks = 0
    i = 0
    while i < len(children):
        ch = children[i]
        is_code_p = False
        if ch.tag == f"{{{W_URI}}}p" and ch.find(".//w:sectPr", NS) is None:
            ppr = ch.find("w:pPr", NS)
            if ppr is not None:
                ps = ppr.find("w:pStyle", NS)
                if ps is not None and ps.attrib.get(W_VAL) == "Code":
                    is_code_p = True
        if not is_code_p:
            i += 1
            continue

        j = i
        block_paras = []
        while j < len(children):
            cj = children[j]
            ok = False
            if cj.tag == f"{{{W_URI}}}p" and cj.find(".//w:sectPr", NS) is None:
                pprj = cj.find("w:pPr", NS)
                if pprj is not None:
                    psj = pprj.find("w:pStyle", NS)
                    if psj is not None and psj.attrib.get(W_VAL) == "Code":
                        ok = True
            if not ok:
                break
            block_paras.append(cj)
            j += 1

        code_lines = ["".join(t.text or "" for t in bp.findall(".//w:t", NS)) for bp in block_paras]
        tbl_el = build_framed_code_table_xml(code_lines, font_half_pt=font_half_pt)
        insert_pos = list(body).index(block_paras[0])
        for bp in block_paras:
            body.remove(bp)
        body.insert(insert_pos, tbl_el)
        wrapped_blocks += 1
        i = j

    return wrapped_blocks


def build_word_table_xml(rows_data: list) -> ET.Element:
    """Tạo bảng Word <w:tbl> với viền đơn chuẩn KMA, hàng đầu dùng Table-Header, các hàng sau dùng Table."""
    tbl = make_w_el("tbl")
    tbl_pr = make_w_el("tblPr")
    tbl_pr.append(make_w_el("tblW", {"w": "0", "type": "auto"}))
    borders = make_w_el("tblBorders")
    for bname in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(make_w_el(bname, {"val": "single", "sz": "4", "space": "0", "color": "000000"}))
    tbl_pr.append(borders)
    tbl.append(tbl_pr)

    for r_idx, row in enumerate(rows_data):
        tr = make_w_el("tr")
        cell_style = "Table-Header" if r_idx == 0 else "Table"
        for cell_text in row:
            tc = make_w_el("tc")
            tc_pr = make_w_el("tcPr")
            tc_pr.append(make_w_el("tcW", {"w": "0", "type": "auto"}))
            tc.append(tc_pr)
            p = build_paragraph_xml(clean_inline_markdown(cell_text.strip()), style_id=cell_style)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    return tbl


def update_docx_xml(docx_path: str, mutator_fn):
    """Đọc word/document.xml và word/styles.xml, gọi mutator_fn(body, auto_numbered_styles), rồi ghi lại file .docx an toàn."""
    tmp_docx = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        doc_root, doc_ns = parse_ooxml_part(zin.read("word/document.xml"))
        styles_root, styles_ns = parse_ooxml_part(zin.read("word/styles.xml"))
        auto_num_styles = ensure_kma_styles_in_docx(styles_root)
        body = doc_root.find("w:body", NS)
        if body is None:
            raise RuntimeError("Không tìm thấy w:body trong word/document.xml")

        mutator_fn(body, auto_num_styles)
        _sort_all_pprs(styles_root)
        _sort_all_pprs(doc_root)

        doc_bytes = serialize_ooxml_part(doc_root, ns_map=doc_ns)
        styles_bytes = serialize_ooxml_part(styles_root, ns_map=styles_ns)

        with zipfile.ZipFile(tmp_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    write_clean_zip_entry(zout, item.filename, doc_bytes, item.date_time)
                elif item.filename == "word/styles.xml":
                    write_clean_zip_entry(zout, item.filename, styles_bytes, item.date_time)
                else:
                    write_clean_zip_entry(zout, item.filename, zin.read(item.filename), item.date_time)
    ensure_docx_writable_and_replace(tmp_docx, docx_path)


def find_insert_index(body: ET.Element, after_heading: str = None) -> int:
    """
    Tìm vị trí chèn trong w:body:
    - Nếu chỉ định after_heading: chèn ngay cuối mục đó (trước đề mục cùng cấp/cấp cao hơn tiếp theo).
    - Nếu không: chèn ngay trước w:sectPr cuối cùng của body.
    """
    children = list(body)
    if after_heading:
        target_norm = after_heading.strip().lower()
        found_heading = False
        for i, ch in enumerate(children):
            if ch.tag == f"{{{W_URI}}}p":
                txt = "".join(t.text for t in ch.findall(".//w:t", NS) if t.text).strip()
                ppr = ch.find("w:pPr", NS)
                st = "Normal"
                if ppr is not None:
                    ps = ppr.find("w:pStyle", NS)
                    if ps is not None:
                        st = ps.attrib.get(W_VAL, "Normal")
                if not found_heading:
                    if target_norm in txt.lower():
                        found_heading = True
                else:
                    if st in ("Heading1", "Heading2", "Heading3", "UH1"):
                        return i
    for i in range(len(children) - 1, -1, -1):
        if children[i].tag == f"{{{W_URI}}}sectPr":
            return i
    return len(children)


def cmd_clone_template(dest_path: str, clean_body: bool = False):
    tpl_path = get_template_path()
    if not os.path.exists(tpl_path):
        print(f"[LỖI] Không tìm thấy file mẫu tại: {tpl_path}")
        sys.exit(1)
    shutil.copy2(tpl_path, dest_path)

    if clean_body:
        # Giữ nguyên 100% Trang bìa, Mục lục, Danh mục đầu bài và Section Break (tới hết child 137),
        # chỉ xóa nội dung mẫu RSA từ Chương 1 trở đi để sẵn sàng ghi báo cáo mới bằng write-md.
        def _clean(body, _auto_num):
            children = list(body)
            second_sect_idx = None
            sect_count = 0
            for idx, ch in enumerate(children):
                if ch.tag == f"{{{W_URI}}}p" and ch.find(".//w:sectPr", NS) is not None:
                    sect_count += 1
                    if sect_count == 2:
                        second_sect_idx = idx
                        break
            if second_sect_idx is not None:
                for ch in children[second_sect_idx + 1:]:
                    if ch.tag != f"{{{W_URI}}}sectPr":
                        body.remove(ch)

        update_docx_xml(dest_path, _clean)
        print(
            f"[OK] Đã tạo file mẫu chuẩn KMA (giữ nguyên Bìa, Mục lục, Section Break và làm sạch thân chương mẫu) tại: {dest_path}"
        )
    else:
        print(f"[OK] Đã tạo bản sao file mẫu chuẩn ĐA KMA tại: {dest_path}")


def cmd_add_placeholder(docx_path: str, fig_num: str, title: str, img_type: str, hint: str, after_heading: str = None):
    if not os.path.exists(docx_path):
        print(f"[LỖI] Không tìm thấy file: {docx_path}")
        sys.exit(1)

    def _insert(body, _auto_num):
        pos = find_insert_index(body, after_heading)
        els = build_image_placeholder_xml(fig_num, title, img_type, hint)
        for offset, el in enumerate(els):
            body.insert(pos + offset, el)

    update_docx_xml(docx_path, _insert)
    print(
        f"[OK] Đã chèn trực tiếp khung giữ chỗ [Hình {fig_num}. {title}] "
        f"(style Hnhv đóng khung + Tnhnhv kèm trường STYLEREF/SEQ) vào: {docx_path}"
    )


def parse_markdown_to_docx_elements(md_text: str, auto_num_styles: set):
    """
    Chuyển đổi nội dung Markdown sang danh sách các phần tử XML Word chuẩn KMA:
    - # Chương X. Tên chương -> Heading1 (hoặc UH1 cho Lời mở đầu, Kết luận, Tài liệu tham khảo, Phụ lục)
    - ## 1.1. Tên mục -> Heading2
    - ### 1.1.1. Tên tiểu mục -> Heading3
    - Khối ```code``` -> style Code
    - Khối giữ chỗ hình ảnh:
        [KHUNG GIỮ CHỖ HÌNH 2.1]
        - Loại hình: ...
        - Nội dung: ...
        - Tên hình: Hình 2.1. ...
      -> Khung giữ chỗ style Hnhv + dòng tên hình style Tnhnhv (có trường STYLEREF/SEQ)
    - Bảng Markdown kèm dòng 'Bảng X.Y. Tên bảng' -> Caption style Tnbng + <w:tbl>
    - Đoạn văn thường -> style Normal (hoặc Tiliuthamkho khi đang ở trong mục Tài liệu tham khảo)
    """
    lines = md_text.splitlines()
    elements = []
    in_code = False
    in_refs = False
    i = 0
    n = len(lines)

    unnumbered_h1 = {
        "lời cảm ơn", "lời nói đầu", "lời mở đầu", "mở đầu", "tóm tắt nội dung đề tài",
        "kết luận", "kết quả và hướng phát triển", "tài liệu tham khảo", "phụ lục",
        "danh mục từ viết tắt", "danh mục kí hiệu và viết tắt", "danh mục hình ảnh",
        "danh mục hình vẽ", "danh mục bảng",
    }

    while i < n:
        raw_line = lines[i]
        s = raw_line.strip()

        # 1. Khối code
        if s.startswith("```"):
            in_code = not in_code
            i += 1
            continue
        if in_code:
            elements.append(build_paragraph_xml(raw_line.rstrip(), style_id="Code"))
            i += 1
            continue

        if not s:
            i += 1
            continue

        # 2. Khối giữ chỗ hình ảnh dạng chuẩn trong Markdown
        ph_match = re.match(
            r"^(?:>\s*)?(?:\*\*)?\[(?:KHUNG\s+GIỮ\s+CHỖ\s+HÌNH|KHUNG\s+HÌNH|ĐỂ\s+TRỐNG\s+HÌNH\s+ẢNH|CHÈN\s+HÌNH)\s+([\d\.]+)(?:[^\]]*)\](?:\*\*)?(.*)$",
            s,
            re.IGNORECASE,
        )
        if ph_match:
            fig_num = ph_match.group(1).strip()
            inline_rest = ph_match.group(2).strip(" :-—*")
            img_type = ""
            hint = inline_rest
            fig_title = f"Hình minh họa {fig_num}"
            i += 1
            while i < n:
                nxt = lines[i].strip()
                if not nxt:
                    i += 1
                    break
                nxt_clean = re.sub(r"^(?:>\s*)?(?:[-*+•]\s*)?(?:\*\*)?", "", nxt)
                if re.match(r"(?i)^loại\s+hình[^:]*:\s*", nxt_clean):
                    img_type = re.sub(r"(?i)^loại\s+hình[^:]*:\s*\**", "", nxt_clean).strip(" *")
                    i += 1
                elif re.match(r"(?i)^(gợi\s+ý|nội\s+dung|mô\s+tả)[^:]*:\s*", nxt_clean):
                    hint = re.sub(r"(?i)^(gợi\s+ý|nội\s+dung|mô\s+tả)[^:]*:\s*\**", "", nxt_clean).strip(" *")
                    i += 1
                elif re.match(r"(?i)^(tên\s+hình|caption|hình\s+\d)[^:]*[:\.]\s*", nxt_clean):
                    fig_title = re.sub(r"(?i)^(tên\s+hình|caption)\s*:\s*\**", "", nxt_clean).strip(" *")
                    i += 1
                else:
                    break
            elements.extend(build_image_placeholder_xml(fig_num, fig_title, img_type, hint))
            continue

        # 3. Đề mục Markdown (#, ##, ###, ####)
        h_match = re.match(r"^(#{1,4})\s+(.+)$", s)
        if h_match:
            level = len(h_match.group(1))
            h_text = clean_inline_markdown(h_match.group(2).strip()).rstrip(".:?;")
            h_low = re.sub(r"^[IVX]+\.\s*", "", h_text).strip().lower()

            if level == 1:
                if h_low in unnumbered_h1:
                    style_id = "UH1"
                    in_refs = "tài liệu tham khảo" in h_low
                    clean_text = re.sub(r"^[IVX]+\.\s*", "", h_text).strip()
                else:
                    style_id = "Heading1"
                    in_refs = False
                    clean_text = h_text
                    if "Heading1" in auto_num_styles:
                        clean_text = re.sub(r"(?i)^chương\s+(?:[0-9]+|[IVX]+)[\.\:]?\s*", "", clean_text).strip()
            elif level == 2:
                style_id = "Heading2"
                clean_text = h_text
                if "Heading2" in auto_num_styles:
                    clean_text = re.sub(r"^\d+\.\d+\.?\s*", "", clean_text).strip()
            elif level == 3:
                style_id = "Heading3"
                clean_text = h_text
                if "Heading3" in auto_num_styles:
                    clean_text = re.sub(r"^\d+\.\d+\.\d+\.?\s*", "", clean_text).strip()
            else:
                style_id = "Heading4"
                clean_text = h_text
                if "Heading4" in auto_num_styles:
                    clean_text = re.sub(r"^\d+(?:\.\d+){3}\.?\s*", "", clean_text).strip()

            elements.append(build_paragraph_xml(clean_text, style_id=style_id))
            i += 1
            continue

        # 4. Caption Bảng (Bảng X.Y. Tên bảng — bắt buộc có dấu '.' hoặc ':' sau số hiệu để không bắt nhầm câu văn mở đầu bằng 'Bảng X.Y tổng hợp...')
        tbl_cap_match = re.match(r"^(?:\*\*)?Bảng\s+(\d+\.\d+)[\.\:]\s+(.+?)(?:\*\*)?$", s, re.IGNORECASE)
        if tbl_cap_match:
            elements.append(build_table_caption_xml(tbl_cap_match.group(1), tbl_cap_match.group(2)))
            i += 1
            continue

        # 5. Bảng Markdown (| Cột 1 | Cột 2 |)
        if s.startswith("|") and s.endswith("|"):
            table_rows = []
            while i < n and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                row_str = lines[i].strip()[1:-1]
                cells = [c.strip() for c in row_str.split("|")]
                if not all(re.match(r"^[:\-]+$", c) for c in cells if c):
                    table_rows.append(cells)
                i += 1
            if table_rows:
                elements.append(build_word_table_xml(table_rows))
            continue

        # 6. Đoạn văn thường hoặc mục Tài liệu tham khảo (gỡ sạch backtick ` và cú pháp Markdown)
        clean_p = clean_inline_markdown(s)
        if in_refs:
            elements.append(build_paragraph_xml(clean_p, style_id="Tiliuthamkho"))
        elif clean_p.startswith(("$ ", "# ", "C:\\>", "PS ")):
            elements.append(build_paragraph_xml(clean_p, style_id="Code"))
        elif re.match(r"^[-*]\s+\S", clean_p):
            bullet_txt = "• " + re.sub(r"^[-*]\s+", "", clean_p)
            elements.append(build_paragraph_xml(bullet_txt, style_id="NoIndent"))
        else:
            elements.append(build_paragraph_xml(clean_p, style_id="Normal"))
        i += 1

    return elements


def _normalize_heading_key(text: str) -> str:
    """Chuẩn hóa tên đề mục (bỏ chữ Chương, số thứ tự đầu mục, dấu câu) để so khớp với khung có sẵn."""
    s = text.strip().lower().rstrip(".:?;")
    s = re.sub(r"^(?:chương\s+(?:[0-9]+|[ivx]+)[\.\:]?\s*)", "", s)
    s = re.sub(r"^(?:[ivx]+\.\s*|\d+(?:\.\d+)*[\.\)]?\s*)", "", s)
    s = re.sub(r"\bchương\s+iii\b", "chương 3", s)
    s = re.sub(r"\bchương\s+ii\b", "chương 2", s)
    s = re.sub(r"\bchương\s+i\b", "chương 1", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _extract_heading_num_prefix(text: str) -> str:
    """Lấy tiền tố đánh số đề mục (ví dụ '1.1', '2.1.3') nếu có."""
    m = re.match(r"^\s*(\d+(?:\.\d+)+)\b", text.strip())
    return m.group(1) if m else ""


def find_existing_heading_in_body(body: ET.Element, raw_heading_text: str, style_id: str):
    """
    Tìm đề mục đã có sẵn trong phần thân của file .docx khung (bỏ qua vùng Mục lục TOC)
    để điền nội dung vào ngay dưới đề mục đó mà không tạo đề mục trùng lặp.
    """
    target_key = _normalize_heading_key(raw_heading_text)
    target_num = _extract_heading_num_prefix(raw_heading_text)
    if not target_key and not target_num:
        return None

    heading_styles = {"Heading1", "Heading2", "Heading3", "Heading4", "UH1", "UH2"}
    children = list(body)
    for idx, ch in enumerate(children):
        if ch.tag != f"{{{W_URI}}}p":
            continue
        ppr = ch.find("w:pPr", NS)
        st = "Normal"
        if ppr is not None:
            ps = ppr.find("w:pStyle", NS)
            if ps is not None:
                st = ps.attrib.get(W_VAL, "Normal")
        if st.startswith("TOC") or st in ("TrangBa", "Mclc", "TableofFigures"):
            continue
        txt = "".join(t.text for t in ch.findall(".//w:t", NS) if t.text).strip()
        if not txt:
            continue
        if st in heading_styles:
            cand_key = _normalize_heading_key(txt)
            cand_num = _extract_heading_num_prefix(txt)
            if target_key and cand_key == target_key:
                return idx
            if target_num and cand_num == target_num and target_key and (target_key in cand_key or cand_key in target_key):
                return idx
    return None


def find_section_end_index(body: ET.Element, heading_idx: int) -> int:
    """Tìm vị trí kết thúc của một đề mục có sẵn (ngay trước đề mục tiếp theo hoặc trước w:sectPr cuối)."""
    heading_styles = {"Heading1", "Heading2", "Heading3", "Heading4", "UH1", "UH2"}
    children = list(body)
    for idx in range(heading_idx + 1, len(children)):
        ch = children[idx]
        if ch.tag == f"{{{W_URI}}}sectPr":
            return idx
        if ch.tag == f"{{{W_URI}}}p":
            ppr = ch.find("w:pPr", NS)
            st = "Normal"
            if ppr is not None:
                ps = ppr.find("w:pStyle", NS)
                if ps is not None:
                    st = ps.attrib.get(W_VAL, "Normal")
            if st in heading_styles:
                return idx
    return len(children)


def cmd_write_md(docx_path: str, md_path: str, after_heading: str = None, source_path: str = None):
    if not os.path.exists(docx_path):
        print(f"[LỖI] Không tìm thấy file .docx đích: {docx_path}")
        sys.exit(1)
    if not os.path.exists(md_path):
        print(f"[LỖI] Không tìm thấy file Markdown nguồn: {md_path}")
        sys.exit(1)

    with open(md_path, "r", encoding="utf-8-sig") as f:
        md_text = f.read()

    inserted_count = [0]
    matched_headings = [0]
    inserted_plain_texts = []

    def _write(body, auto_num_styles):
        els = parse_markdown_to_docx_elements(md_text, auto_num_styles)
        for el in els:
            txt_el = "".join(t.text for t in el.findall(".//w:t", NS) if t.text).strip()
            if txt_el and not PLACEHOLDER_MARKER_RE.search(txt_el):
                inserted_plain_texts.append(txt_el)
        if after_heading:
            pos = find_insert_index(body, after_heading)
            for offset, el in enumerate(els):
                body.insert(pos + offset, el)
            inserted_count[0] = len(els)
            return

        # Tách danh sách phần tử thành các cụm (heading_el_hoặc_None, [body_els])
        heading_styles = {"Heading1", "Heading2", "Heading3", "Heading4", "UH1", "UH2"}
        sections = []
        cur_h = None
        cur_h_txt = ""
        cur_h_style = ""
        cur_items = []

        for el in els:
            st = "Normal"
            if el.tag == f"{{{W_URI}}}p":
                ppr = el.find("w:pPr", NS)
                if ppr is not None:
                    ps = ppr.find("w:pStyle", NS)
                    if ps is not None:
                        st = ps.attrib.get(W_VAL, "Normal")
            if st in heading_styles:
                if cur_h is not None or cur_items:
                    sections.append((cur_h, cur_h_txt, cur_h_style, cur_items))
                cur_h = el
                cur_h_txt = "".join(t.text for t in el.findall(".//w:t", NS) if t.text).strip()
                cur_h_style = st
                cur_items = []
            else:
                cur_items.append(el)
        if cur_h is not None or cur_items:
            sections.append((cur_h, cur_h_txt, cur_h_style, cur_items))

        for h_el, h_txt, h_st, body_els in sections:
            if h_el is not None:
                existing_idx = find_existing_heading_in_body(body, h_txt, h_st)
                if existing_idx is not None:
                    # Đề mục đã có sẵn trong khung .docx -> giữ nguyên đề mục gốc, chỉ điền nội dung vào dưới nó
                    matched_headings[0] += 1
                    insert_pos = find_section_end_index(body, existing_idx)
                    for offset, bel in enumerate(body_els):
                        body.insert(insert_pos + offset, bel)
                    inserted_count[0] += len(body_els)
                    continue
            # Đề mục chưa có trong khung .docx -> chèn cả đề mục và nội dung vào cuối phần thân
            insert_pos = find_insert_index(body, None)
            all_to_insert = ([h_el] if h_el is not None else []) + body_els
            for offset, bel in enumerate(all_to_insert):
                body.insert(insert_pos + offset, bel)
            inserted_count[0] += len(all_to_insert)

    update_docx_xml(docx_path, _write)
    print(
        f"[OK] Đã điền trực tiếp {inserted_count[0]} phần tử nội dung (khớp {matched_headings[0]} đề mục có sẵn trong khung, "
        f"bảo toàn đề mục gốc & chèn khung giữ chỗ hình Hnhv/Tnhnhv) từ '{md_path}' vào '{docx_path}'."
    )
    # Tự kiểm tra dữ kiện không bị rớt trong quá trình chuyển Markdown -> OOXML
    md_prose = "\n".join(
        ln for ln in md_text.splitlines()
        if not ln.strip().startswith("```") and not PLACEHOLDER_MARKER_RE.search(ln)
    )
    written_prose = "\n".join(inserted_plain_texts)
    conv_warns = _compare_facts_vi(md_prose, written_prose, check_drops=True)
    for w in conv_warns:
        print("  [!]", w)
    if source_path and os.path.exists(source_path):
        src_items, _ = load_items_from_file(source_path)
        src_prose = "\n".join(t for _, st, t, _ in src_items if st != "Code" and not PLACEHOLDER_MARKER_RE.search(t))
        for w in _compare_facts_vi(src_prose, written_prose):
            print("  [!]", w)


# ==============================================================================
# 5. CLI ENTRYPOINT
# ==============================================================================
def build_arg_parser():
    parser = argparse.ArgumentParser(
        description="Công cụ kiểm tra & ghi trực tiếp file .docx báo cáo BTL / Đồ án tốt nghiệp KMA + Humanizer Tiếng Việt"
    )
    sub = parser.add_subparsers(dest="action")

    p_ins = sub.add_parser("inspect", help="Xem cây đề mục, căn lề, thống kê ViDetect và danh sách khung hình trong .docx")
    p_ins.add_argument("target", help="Đường dẫn file .docx")

    p_chk = sub.add_parser("check", help="Quét lỗi đề mục, cấu trúc OOXML Word, văn dịch sâu, nhịp câu CV và mẫu câu AI tiếng Việt")
    p_chk.add_argument("target", help="Đường dẫn file .docx, .md hoặc .pdf")
    p_chk.add_argument("--mode", choices=["do-an", "btl"], default="do-an", help="Chế độ kiểm tra: do-an (ĐATN) hoặc btl (Bài tập lớn)")
    p_chk.add_argument("--no-cite", action="store_true", default=True, help="Bỏ qua kiểm tra trích dẫn [n] trong thân bài (mặc định bật)")
    p_chk.add_argument("--require-cite", action="store_true", help="Bắt buộc kiểm tra trích dẫn [1], [2]... trong thân bài")
    p_chk.add_argument("--source", default=None, help="File nguồn/bản gốc (.pdf, .docx, .md) để đối chiếu không mất/sai số liệu và mức khẳng định P4")

    p_scr = sub.add_parser("score-local", help="Tiền xử lý chuẩn VietAIDetector, thống kê ViDetect và chấm thử với PhoBERT cục bộ (nếu có)")
    p_scr.add_argument("target", help="Đường dẫn file .docx, .md hoặc .pdf")

    p_cln = sub.add_parser("clone-template", help="Tạo file .docx mới từ bản mẫu gốc Mau_trinh_bay_DA.docx")
    p_cln.add_argument("dest", help="Đường dẫn file .docx mới")
    p_cln.add_argument(
        "--clean-body",
        action="store_true",
        help="Giữ nguyên trang bìa, mục lục, section breaks nhưng xóa phần nội dung mẫu RSA để ghi bài mới",
    )

    p_ph = sub.add_parser("add-placeholder", help="Ghi trực tiếp khối giữ chỗ hình ảnh đóng khung (style Hnhv + Tnhnhv) vào file .docx")
    p_ph.add_argument("docx", help="Đường dẫn file .docx")
    p_ph.add_argument("--fig", required=True, help="Số hiệu hình (ví dụ: 2.1)")
    p_ph.add_argument("--title", required=True, help="Tên hình đặt dưới khung (ví dụ: Sơ đồ tấn công tràn băng thông)")
    p_ph.add_argument("--type", default="Sơ đồ khối / Ảnh chụp màn hình thực nghiệm", help="Loại hình gợi ý")
    p_ph.add_argument("--hint", required=True, help="Gợi ý chi tiết các thành phần cần có trong hình để tự thêm")
    p_ph.add_argument("--after-heading", default=None, help="Tên đề mục muốn chèn vào sau đó (tùy chọn)")

    p_wmd = sub.add_parser("write-md", help="Ghi trực tiếp nội dung từ file Markdown (.md) vào file .docx theo đúng bộ Style KMA")
    p_wmd.add_argument("docx", help="Đường dẫn file .docx đích")
    p_wmd.add_argument("md", help="Đường dẫn file .md nguồn")
    p_wmd.add_argument("--after-heading", default=None, help="Chèn vào cuối đề mục chỉ định (mặc định chèn vào cuối phần nội dung)")
    p_wmd.add_argument("--apply-format", action="store_true", help="Áp dụng luôn định dạng trang A4, lề, chân trang, căn chỉnh ảnh và bộ Style chuẩn ĐATN")
    p_wmd.add_argument("--source", default=None, help="File nguồn tham chiếu để kiểm tra giữ đúng số liệu và mức khẳng định P4 khi ghi")

    p_fmt = sub.add_parser("apply-format", help="Chuẩn hóa định dạng trang A4, lề, chân trang, font, căn chỉnh ảnh và bộ Style ĐATN KMA")
    p_fmt.add_argument("docx", help="Đường dẫn file .docx cần chuẩn hóa định dạng")
    p_fmt.add_argument("--max-img-width-cm", type=float, default=15.5, help="Chiều rộng tối đa của ảnh trong thân bài (cm)")
    p_fmt.add_argument("--max-img-height-cm", type=float, default=11.0, help="Chiều cao tối đa của ảnh trong thân bài (cm)")

    p_ali = sub.add_parser("align-images", help="Căn giữa style Hnhv và co giãn tỷ lệ toàn bộ hình ảnh trong thân bài .docx")
    p_ali.add_argument("docx", help="Đường dẫn file .docx cần căn chỉnh ảnh")
    p_ali.add_argument("--max-width-cm", type=float, default=15.5, help="Chiều rộng tối đa của ảnh (cm, mặc định 15.5)")
    p_ali.add_argument("--max-height-cm", type=float, default=11.0, help="Chiều cao tối đa của ảnh (cm, mặc định 11.0)")

    p_val = sub.add_parser("validate-xml", help="Kiểm tra toàn vẹn cấu trúc OOXML để đảm bảo mở/lưu trực tiếp trong Word không lỗi")
    p_val.add_argument("docx", help="Đường dẫn file .docx cần kiểm tra")

    return parser


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = build_arg_parser()
    if len(sys.argv) < 2:
        parser.print_help()
        sys.exit(0)
    args = parser.parse_args()
    try:
        if args.action == "inspect":
            cmd_inspect(args.target)
        elif args.action == "check":
            cmd_check(
                args.target,
                mode=args.mode,
                no_cite=args.no_cite,
                require_cite=args.require_cite,
                source_path=args.source,
            )
        elif args.action == "score-local":
            cmd_score_local(args.target)
        elif args.action == "clone-template":
            cmd_clone_template(args.dest, clean_body=args.clean_body)
        elif args.action == "add-placeholder":
            cmd_add_placeholder(
                args.docx,
                fig_num=args.fig,
                title=args.title,
                img_type=args.type,
                hint=args.hint,
                after_heading=args.after_heading,
            )
        elif args.action == "write-md":
            cmd_write_md(args.docx, args.md, after_heading=args.after_heading, source_path=args.source)
            if args.apply_format:
                cmd_apply_format(args.docx)
        elif args.action == "apply-format":
            cmd_apply_format(
                args.docx,
                max_img_width_cm=args.max_img_width_cm,
                max_img_height_cm=args.max_img_height_cm,
            )
        elif args.action == "align-images":
            cmd_align_images(
                args.docx,
                max_width_cm=args.max_width_cm,
                max_height_cm=args.max_height_cm,
            )
        elif args.action == "validate-xml":
            probs = validate_docx_ooxml(args.docx)
            if not probs:
                print(f"[ĐẠT OOXML] Cấu trúc XML hợp lệ 100%, mở và Ctrl+S trực tiếp trong Word an toàn: {args.docx}")
            else:
                print(f"[LỖI OOXML] Phát hiện {len(probs)} vấn đề:")
                for pr in probs:
                    print("  *", pr)
                sys.exit(1)
        else:
            parser.print_help()
            sys.exit(1)
    except BrokenPipeError:
        pass
