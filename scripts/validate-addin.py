#!/usr/bin/env python3
"""
Kiem tra mot file LINK.xlam co du dieu kien phat hanh khong.

Vi sao can: mot ban .xlam co the trong "binh thuong" nhung da mat kha nang tu
cap nhat - vi du file duoc xuat tu ban .xlsm goc chua duoc dan doan goi trong
ThisWorkbook. Neu phat hanh ban do, MOI may se cap nhat mot lan len no roi mat
kha nang tu cap nhat vinh vien, phai di cai lai tung may bang install.bat.

Kiem tra bang cach doc that su ma nguon VBA: vbaProject.bin la mot OLE compound
file, ma nguon moi module nam trong do o dang nen RLE cua MS-OVBA. Tim chuoi
tho tren file nhi phan KHONG dang tin - da thu va cho ket qua sai - nen o day
doc dung cau truc roi giai nen.

Dung:  python3 validate-addin.py <duong-dan.xlam>
Ma thoat: 0 = dat, 1 = khong dat, 2 = khong doc duoc file.
"""

import re
import struct
import sys
import zipfile

ADDIN_CONTENT_TYPE = "application/vnd.ms-excel.addin.macroEnabled.main+xml"
EXPECTED_RAW_BASE = "https://raw.githubusercontent.com/namtao/add-in/main/release/"
MIN_BYTES = 51200

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF


class CompoundFile:
    """Doc OLE compound file (vbaProject.bin) - chi phan can de lay stream."""

    def __init__(self, data):
        if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("khong phai OLE compound file")
        self.data = data
        self.sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", data, 0x20)[0]
        first_dir = struct.unpack_from("<I", data, 0x30)[0]
        self.mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
        first_mini_fat = struct.unpack_from("<I", data, 0x3C)[0]
        first_difat = struct.unpack_from("<I", data, 0x44)[0]
        num_difat = struct.unpack_from("<I", data, 0x48)[0]

        difat = list(struct.unpack_from("<109I", data, 0x4C))
        sector = first_difat
        for _ in range(num_difat):
            if sector in (ENDOFCHAIN, FREESECT):
                break
            raw = self._sector(sector)
            per = self.sector_size // 4 - 1
            difat.extend(struct.unpack_from("<%dI" % per, raw, 0))
            sector = struct.unpack_from("<I", raw, per * 4)[0]

        self.fat = []
        for s in difat:
            if s in (ENDOFCHAIN, FREESECT):
                continue
            raw = self._sector(s)
            self.fat.extend(
                struct.unpack_from("<%dI" % (self.sector_size // 4), raw, 0)
            )

        self.mini_fat = []
        sector = first_mini_fat
        while sector not in (ENDOFCHAIN, FREESECT):
            raw = self._sector(sector)
            self.mini_fat.extend(
                struct.unpack_from("<%dI" % (self.sector_size // 4), raw, 0)
            )
            sector = self.fat[sector]

        self.entries = self._read_directory(first_dir)
        root = self.entries[0]
        self.mini_stream = self._read_chain(root["start"], root["size"])

    def _sector(self, n):
        off = (n + 1) * self.sector_size
        return self.data[off : off + self.sector_size]

    def _read_chain(self, start, size):
        out = bytearray()
        sector = start
        while sector not in (ENDOFCHAIN, FREESECT) and len(out) < size:
            out.extend(self._sector(sector))
            sector = self.fat[sector]
        return bytes(out[:size])

    def _read_mini_chain(self, start, size):
        out = bytearray()
        sector = start
        while sector not in (ENDOFCHAIN, FREESECT) and len(out) < size:
            off = sector * self.mini_sector_size
            out.extend(self.mini_stream[off : off + self.mini_sector_size])
            sector = self.mini_fat[sector]
        return bytes(out[:size])

    def _read_directory(self, first_dir):
        raw = bytearray()
        sector = first_dir
        while sector not in (ENDOFCHAIN, FREESECT):
            raw.extend(self._sector(sector))
            sector = self.fat[sector]
        entries = []
        for off in range(0, len(raw), 128):
            chunk = raw[off : off + 128]
            if len(chunk) < 128:
                break
            name_len = struct.unpack_from("<H", chunk, 64)[0]
            if name_len < 2:
                continue
            name = chunk[: name_len - 2].decode("utf-16-le", "replace")
            entries.append(
                {
                    "name": name,
                    "type": chunk[66],
                    "start": struct.unpack_from("<I", chunk, 116)[0],
                    "size": struct.unpack_from("<Q", chunk, 120)[0],
                }
            )
        return entries

    def stream(self, name):
        """Lay noi dung stream theo ten. Ten trong vbaProject.bin la duy nhat."""
        for e in self.entries:
            if e["name"] == name and e["type"] == 2:
                if e["size"] < self.mini_cutoff:
                    return self._read_mini_chain(e["start"], e["size"])
                return self._read_chain(e["start"], e["size"])
        return None


def decompress(data, start=0):
    """Giai nen container theo MS-OVBA 2.4.1.3 (RLE)."""
    if start >= len(data) or data[start] != 0x01:
        raise ValueError("thieu chu ky 0x01 cua vung nen")
    out = bytearray()
    i = start + 1
    while i + 1 < len(data):
        header = struct.unpack_from("<H", data, i)[0]
        i += 2
        size = (header & 0x0FFF) + 3
        compressed = (header & 0x8000) != 0
        end = min(i + size - 2, len(data))
        if not compressed:
            out.extend(data[i:end])
            i = end
            continue
        chunk_start = len(out)
        while i < end:
            flags = data[i]
            i += 1
            for bit in range(8):
                if i >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[i])
                    i += 1
                else:
                    token = struct.unpack_from("<H", data, i)[0]
                    i += 2
                    diff = len(out) - chunk_start
                    bit_count = 4
                    while (1 << bit_count) < diff:
                        bit_count += 1
                    length = (token & (0xFFFF >> bit_count)) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    src = len(out) - offset
                    for k in range(length):
                        out.append(out[src + k])
    return bytes(out)


def module_text_offset(dir_text, module_name):
    """Tim TextOffset cua mot module trong stream 'dir' da giai nen.

    Khong duyet toan bo ban ghi tu dau: vai ban ghi REFERENCE* co cau truc long
    nhau lam lech con tro. Thay vao do tim thang ban ghi MODULENAME (id 0x0019)
    dung ten can tim, roi quet ve sau tim MODULEOFFSET (id 0x0031) gan nhat.
    """
    name = module_name.encode("latin-1")
    needle = struct.pack("<HI", 0x0019, len(name)) + name
    pos = dir_text.find(needle)
    if pos < 0:
        return None
    window = dir_text[pos : pos + 512]
    m = window.find(struct.pack("<HI", 0x0031, 4))
    if m < 0:
        return None
    return struct.unpack_from("<I", window, m + 6)[0]


def module_source(cfb, dir_text, module_name):
    raw = cfb.stream(module_name)
    if raw is None:
        return None
    offset = module_text_offset(dir_text, module_name)
    if offset is None or offset >= len(raw):
        return None
    return decompress(raw, offset).decode("latin-1")


def strip_comments(src):
    """Bo chu thich VBA, nhung giu nguyen dau nhay don nam ben trong chuoi.

    Khong duoc cat tho tai dau ' dau tien cua dong: ma nguon co nhung dong nhu
    "'" & ThisWorkbook.Name & "'!modAutoUpdate.CheckForUpdate"
    va cat o do se vut di dung phan can kiem tra. Vi vay phai bam theo trang
    thai dang-o-trong-chuoi. Trong VBA khong co ky tu escape: dau " doi dang
    moi lan gap, con "" long nhau doi hai lan nen ket qua van dung.
    """
    out = []
    for line in src.splitlines():
        in_str = False
        cut = None
        for i, ch in enumerate(line):
            if ch == '"':
                in_str = not in_str
            elif ch == "'" and not in_str:
                cut = i
                break
        code = line if cut is None else line[:cut]
        if code.strip().lower().startswith("rem "):
            continue
        out.append(code)
    return "\n".join(out)


def validate(path):
    """Tra ve danh sach loi. Rong = dat."""
    errors = []

    with open(path, "rb") as fh:
        raw = fh.read()
    if len(raw) < MIN_BYTES:
        return ["File chi co %d bytes - qua nho, khong phai LINK.xlam." % len(raw)]
    if raw[:2] != b"PK":
        return ["File khong phai dinh dang .xlam hop le (thieu chu ky ZIP 'PK')."]

    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return ["File ZIP hong, khong doc duoc."]

    names = set(zf.namelist())

    content_types = zf.read("[Content_Types].xml").decode("utf-8", "replace")
    if ADDIN_CONTENT_TYPE not in content_types:
        errors.append(
            "Day KHONG phai file add-in (.xlam). Co ve ban dang phat hanh mot "
            "workbook .xlsm. Trong VBE dat ThisWorkbook.IsAddin = True roi Save As "
            "sang .xlam."
        )

    workbook = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    view = re.search(r"<workbookView[^>]*>", workbook)
    if not view or 'visibility="veryHidden"' not in view.group(0):
        errors.append(
            "File duoc luu luc IsAddin = False nen cua so workbook se hien ra khi "
            "nap. Trong VBE chon ThisWorkbook, bam F4, dat IsAddin = True roi luu lai."
        )

    if "xl/vbaProject.bin" not in names:
        errors.append("File khong chua ma VBA nao (thieu xl/vbaProject.bin).")
        return errors

    vba = zf.read("xl/vbaProject.bin")
    try:
        cfb = CompoundFile(vba)
        dir_text = decompress(cfb.stream("dir"))
    except Exception as exc:
        errors.append("Khong doc duoc ma VBA trong file: %s" % exc)
        return errors

    auto = module_source(cfb, dir_text, "modAutoUpdate")
    if auto is None:
        errors.append(
            "File CHUA co module tu cap nhat (modAutoUpdate). Hay import "
            "src/modAutoUpdate.bas vao file goc."
        )
    else:
        body = strip_comments(auto)
        if "Sub CheckForUpdate" not in body:
            errors.append("Module modAutoUpdate co nhung thieu Sub CheckForUpdate.")
        # Application.OnTime phan giai ten thu tuc theo workbook dang active luc
        # timer ban, khong theo workbook goi no. Thieu tien to "'<ten file>'!"
        # thi lich hen bi bo qua trong im lang va bo tu cap nhat khong bao gio
        # chay - dung loi da lam ban cu chet lang. Chan tai day.
        if not re.search(
            r"Application\.OnTime[\s\S]{0,400}?'!modAutoUpdate\.CheckForUpdate", body
        ):
            errors.append(
                "Loi goi Application.OnTime trong modAutoUpdate thieu tien to ten "
                "workbook, nen CheckForUpdate se khong bao gio chay. Ten macro phai "
                'co dang "\'" & ThisWorkbook.Name & "\'!modAutoUpdate.CheckForUpdate".'
            )
        # Trim$ trong VBA khong cat CR/LF, ma version.txt luon ket thuc bang
        # newline. Ban dung Trim$ de lam sach checksum se so sanh sai o MOI phep
        # StrComp va khong bao gio cap nhat duoc, hoan toan im lang. CleanHash
        # la ham loc lay dung ky tu hex, thay cho Trim$.
        if "CleanHash" not in body:
            errors.append(
                "modAutoUpdate khong co CleanHash de lam sach checksum. Dung Trim$ "
                "la sai vi no khong cat CR/LF, ma version.txt luon co newline o "
                "cuoi - moi phep so sanh checksum se that bai trong im lang."
            )
        if EXPECTED_RAW_BASE not in auto:
            errors.append(
                "modAutoUpdate dang tro toi mot dia chi khac, khong phai repo nay "
                "(%s)." % EXPECTED_RAW_BASE
            )

    this_wb = module_source(cfb, dir_text, "ThisWorkbook")
    if this_wb is None:
        errors.append("Khong doc duoc module ThisWorkbook.")
    else:
        body = strip_comments(this_wb)
        has_open = re.search(r"\bSub\s+Workbook_Open\s*\(", body, re.IGNORECASE)
        has_call = re.search(r"\bScheduleUpdateCheck\b", body, re.IGNORECASE)
        if not (has_open and has_call):
            errors.append(
                "ThisWorkbook KHONG goi modAutoUpdate.ScheduleUpdateCheck, nen bo tu "
                "cap nhat se khong bao gio chay.\n"
                "        Mo file goc, bam Alt+F11, double-click ThisWorkbook va dan vao:\n"
                "            Private Sub Workbook_Open()\n"
                "                modAutoUpdate.ScheduleUpdateCheck\n"
                "            End Sub\n"
                "        Lam mot lan tren file goc la xong, moi ban xuat ra sau deu co san."
            )

    return errors


def main():
    if len(sys.argv) != 2:
        print("Dung: validate-addin.py <duong-dan.xlam>", file=sys.stderr)
        return 2
    try:
        errors = validate(sys.argv[1])
    except OSError as exc:
        print("Khong doc duoc file: %s" % exc, file=sys.stderr)
        return 2
    if errors:
        print("KHONG DAT - file nay khong duoc phep phat hanh:\n")
        for e in errors:
            print("  - %s" % e)
        return 1
    print("DAT - file hop le, du dieu kien phat hanh.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
