#!

# pip install pywin32

import win32com.client as win32
from pathlib import Path
from time import sleep
from enum import IntEnum
from os import PathLike


class HwpClearOption(IntEnum):
    # https://www.hancom.com/board/devmanualList.do?artcl_seq=3985
    # 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다.
    hwpAskSave = 0
    # 문서의 내용을 버린다.
    hwpDiscard = 1
    # 문서가 변경된 경우 저장한다.
    hwpSaveIfDirty = 2
    # 무조건 저장한다.
    hwpSave = 3


def hwp_save_as_pdf(hwp, hwppath: PathLike, pdfpath: PathLike) -> None:
    try:
        # https://www.hancom.com/board/devmanualList.do?artcl_seq=3997
        # "forceopen:true"
        hwp.Open(hwppath)

        # https://www.hancom.com/board/devmanualList.do?artcl_seq=3916
        ac = hwp.CreateAction("FileSaveAsPdf")
        # https://www.hancom.com/board/devmanualList.do?artcl_seq=3929
        ps = ac.CreateSet()
        ac.GetDefault(ps)

        ps.SetItem("Attributes", 0)
        ps.SetItem("FileName", str(pdfpath))
        ps.SetItem("Format", "PDF")

        ac.Execute(ps)
        print(pdfpath, "done")
    except Exception as e:
        print(hwppath)
        print(e)
    finally:
        sleep(1)
        hwp.Clear(HwpClearOption.hwpDiscard)


def get_hwp_paths(parent: PathLike) -> list[Path]:
    hwpexts = [x.casefold() for x in [".hwp", ".hwpx"]]

    p = Path(parent)
    paths = []
    for child in p.glob("**/*.hwp*"):
        if child.suffix.casefold() in hwpexts:
            paths.append(child)

    return paths


parent = "C:\\Users\\USER\Documents\\auri 연구보고서"
hwppaths = get_hwp_paths(parent)

badfiles = ["[AURI-기본-2013-3] 가로단위 공간관리 수단으로서의 특별가로구역 제도 연구.hwp"]

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
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
