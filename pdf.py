from time import sleep

from win32hwp import get_hwp_object, get_hwp_paths, hwp_save_as_pdf

parent = "C:\\Users\\USER\Documents\\auri 연구보고서"
hwppaths = get_hwp_paths(parent)

badfiles = ["[AURI-기본-2013-3] 가로단위 공간관리 수단으로서의 특별가로구역 제도 연구.hwp"]

hwp = get_hwp_object()
for hp in hwppaths:
    pdfpath = hp.with_suffix(".pdf")

    if hp.name in badfiles:
        print(hp, "skipped")
    elif pdfpath.exists():
        print(pdfpath, "already exists")
    else:
        hwp_save_as_pdf(hwp, hp, pdfpath)

sleep(5)
hwp.Quit()
