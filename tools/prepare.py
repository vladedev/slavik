#!/usr/bin/env python3
"""Подготовка PDF к распознаванию: дубли, цифровые страницы, рендер.

python tools/prepare.py input/Файл.pdf [ещё.pdf ...] [--keep first|last|all]

  first (по умолчанию) — если в конце PDF приклеена повторная копия документа,
                          берём первую версию, копию пропускаем целиком
  last                 — берём только повторную копию (например, чистый оригинал)
  all                  — все уникальные страницы

Результат: work/<имя>/plan.json, work/<имя>/pages/pNNN.jpg,
           work/<имя>/text/pNNN.txt (для страниц с текстовым слоем).
Модель не нужна. Повторный запуск ничего не перерисовывает.
"""
import argparse, json, sys
from pathlib import Path

try:
    import pypdfium2 as pdfium
except ImportError:
    sys.exit("Нужен pypdfium2: pip install pypdfium2 pillow")
from PIL import Image

DUP_DIST = 10        # порог перцептивного хэша (256 бит); дубли дают 0, разные страницы 35+
TEXT_MIN_CHARS = 200 # больше — считаем, что у страницы есть текстовый слой
DPI = 150
MAX_SIDE = 1568      # больше модель всё равно ужмёт


def dhash(img):
    g = img.convert("L").resize((17, 16), Image.LANCZOS)
    px = list(g.tobytes())
    bits = 0
    for r in range(16):
        for c in range(16):
            bits = (bits << 1) | (px[r * 17 + c] > px[r * 17 + c + 1])
    return bits


def find_trailing_copy(dup_of, n):
    """Первая страница k>1, с которой начинается повторная копия документа."""
    for k in range(2, n + 1):
        if dup_of.get(k) is None:
            continue
        tail = range(k, n + 1)
        share = sum(1 for p in tail if dup_of.get(p) is not None) / len(tail)
        if len(tail) >= 3 and share >= 0.6 and dup_of[k] < k:
            return k
    return None


def prepare(pdf_path, keep):
    pdf_path = Path(pdf_path)
    work = Path("work") / pdf_path.stem
    (work / "pages").mkdir(parents=True, exist_ok=True)
    (work / "text").mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(str(pdf_path))
    n = len(pdf)
    hashes, text_pages = {}, []
    for i in range(n):
        p = i + 1
        page = pdf[i]
        txt = page.get_textpage().get_text_range() or ""
        if len(txt.strip()) >= TEXT_MIN_CHARS:
            text_pages.append(p)
            (work / "text" / f"p{p:03d}.txt").write_text(txt, encoding="utf-8")
        hashes[p] = dhash(page.render(scale=40 / 72).to_pil())

    dup_of = {}
    for p in range(2, n + 1):
        for q in range(1, p):
            if bin(hashes[p] ^ hashes[q]).count("1") <= DUP_DIST:
                dup_of[p] = q
                break

    k = find_trailing_copy(dup_of, n)
    skipped = {}
    if keep == "first" and k:
        candidates = list(range(1, k))
        for p in range(k, n + 1):
            skipped[p] = f"повторная копия документа (с стр. {k})"
    elif keep == "last" and k:
        candidates = list(range(k, n + 1))
        for p in range(1, k):
            skipped[p] = f"первая версия (копия начинается со стр. {k})"
    else:
        candidates = list(range(1, n + 1))

    selected = []
    for p in candidates:
        q = dup_of.get(p)
        if q is not None and q in selected:
            skipped[p] = f"дубль стр. {q}"
        else:
            selected.append(p)

    for old in (work / "pages").glob("p*.jpg"):   # убрать страницы от прошлого режима
        if int(old.stem[1:]) not in selected:
            old.unlink()
    for p in selected:
        out = work / "pages" / f"p{p:03d}.jpg"
        if out.exists():
            continue
        img = pdf[p - 1].render(scale=DPI / 72).to_pil().convert("RGB")
        img.thumbnail((MAX_SIDE, MAX_SIDE), Image.LANCZOS)
        img.save(out, "JPEG", quality=85)

    plan = {
        "pdf": str(pdf_path),
        "pages_total": n,
        "keep": keep,
        "trailing_copy_starts_at": k,
        "selected": selected,
        "text_layer_pages": [p for p in text_pages if p in selected],
        "skipped": {str(p): r for p, r in sorted(skipped.items())},
    }
    (work / "plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n== {pdf_path.name}")
    print(f"   страниц в PDF: {n}; к распознаванию: {len(selected)}; пропуск: {len(skipped)}")
    if k:
        print(f"   найдена повторная копия документа со стр. {k} (режим --keep {keep})")
    if plan["text_layer_pages"]:
        print(f"   с текстовым слоем: {plan['text_layer_pages']}")
    if skipped:
        print(f"   пропущены: {compress(sorted(skipped))}")
    print(f"   план: {work / 'plan.json'}")


def compress(nums):
    out, start, prev = [], None, None
    for x in nums:
        if start is None:
            start = prev = x
        elif x == prev + 1:
            prev = x
        else:
            out.append(f"{start}-{prev}" if start != prev else str(start))
            start = prev = x
    if start is not None:
        out.append(f"{start}-{prev}" if start != prev else str(start))
    return ", ".join(out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pdfs", nargs="+")
    ap.add_argument("--keep", choices=["first", "last", "all"], default="first")
    a = ap.parse_args()
    for f in a.pdfs:
        prepare(f, a.keep)
