from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from openpyxl import Workbook

from excel_toolkit.application.preview import PreviewService
from excel_toolkit.infrastructure.files.reader import FilePreviewReader


@pytest.fixture
def service() -> PreviewService:
    return PreviewService(FilePreviewReader())


@pytest.fixture
def xlsx_file(tmp_path: Path) -> Path:
    path = tmp_path / "Клиенты.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Клиенты"
    sheet.append(["Код", "Имя", "Количество", "Формула", "Дата", "Флаг"])
    sheet.append(["00123", "Алматы", 12, "=C2*2", datetime(2026, 1, 2), True])
    sheet["C2"].number_format = "00000"
    second = workbook.create_sheet("Заказы")
    second.append(["00007", "Заказ"])
    workbook.create_sheet("Пустой")
    workbook.save(path)
    workbook.close()
    return path


@pytest.fixture
def feature_bag_xlsx(tmp_path: Path) -> Path:
    """Real OPC package with checkbox style metadata, not a mocked storage adapter."""
    path = tmp_path / "feature-bag.xlsx"
    book = Workbook()
    book.active.append([" h "])
    book.active.append([True])
    book.save(path)
    book.close()
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    namespace = "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"
    parts["xl/featurePropertyBag/featurePropertyBag.xml"] = (
        f'<FeaturePropertyBags xmlns="{namespace}" count="4">'
        '<bag type="Checkbox"/>'
        '<bag type="XFControls"><bagId k="CellControl">0</bagId></bag>'
        '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>'
        '<bag type="XFComplements" extRef="XFComplementsMapperExtRef">'
        '<a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag>'
        "</FeaturePropertyBags>"
    ).encode()
    rel_name = "xl/_rels/workbook.xml.rels"
    relations = ET.fromstring(parts[rel_name])
    ET.SubElement(
        relations,
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
        Id="rIdFeatureBag",
        Type="http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag",
        Target="featurePropertyBag/featurePropertyBag.xml",
    )
    parts[rel_name] = ET.tostring(relations)
    content_types = ET.fromstring(parts["[Content_Types].xml"])
    ET.SubElement(
        content_types,
        "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
        PartName="/xl/featurePropertyBag/featurePropertyBag.xml",
        ContentType="application/vnd.ms-excel.featurepropertybag+xml",
    )
    parts["[Content_Types].xml"] = ET.tostring(content_types)
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    styles = ET.fromstring(parts["xl/styles.xml"])
    xfs = styles.find(f"{{{main}}}cellXfs")
    # Keep the header's style ordinary; only the boolean cell uses the new style.
    xf = ET.SubElement(xfs, f"{{{main}}}xf", dict(xfs[0].attrib))
    xfs.set("count", str(len(xfs)))
    ext_list = ET.SubElement(xf, f"{{{main}}}extLst")
    extension = ET.SubElement(
        ext_list, f"{{{main}}}ext", uri="{C7286773-470A-42A8-94C5-96B5CB345126}"
    )
    ET.SubElement(extension, f"{{{namespace}}}xfComplement", i="0")
    parts["xl/styles.xml"] = ET.tostring(styles)
    sheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    sheet.find(f'.//{{{main}}}c[@r="A2"]').set("s", "1")
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(sheet)
    with ZipFile(path, "w") as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return path
