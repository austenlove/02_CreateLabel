from pathlib import Path
import zipfile
import json
import math
from copy import deepcopy
from collections import OrderedDict
from lxml import etree


# 차수별로 데이터를 묶기 (입력 순서 유지)
def _group_by_round(data):
    # 데이터를 묶을 dictionary
    groups = OrderedDict()

    for item in data:
        round_no = str(item.get("차수", "")).strip()

        if not round_no:
            raise ValueError("차수가 비어 있는 데이터가 있습니다.")

        groups.setdefault(round_no, []).append(item)

    # 차수가 숫자라면 숫자 순으로 정렬
    try:
        groups = OrderedDict(sorted(groups.items(), key=lambda x: int(x[0])))

    except ValueError:
        pass

    return groups



# 주소 카드 셀의 placeholder를 데이터로 치환
def _replace_in_cell(cell, item):

    replacements = {
        "{{주소}}": str(item.get("주소", "")),
        "{{보호자 성명}}": str(item.get("보호자 성명", "")),
        "{{학생 성명}}": str(item.get("학생 성명", "")),
        "{{우편번호}}": str(item.get("우편번호", "")),
    }

    # cell 내부 xml 태그가 t인 모든 텍스트 노드 찾기
    text_nodes = cell.xpath('.//*[local-name()="t"]')

    for node in text_nodes:
        if node.text is None:
            continue

        text = node.text

        for old, new in replacements.items():
            text = text.replace(old, new)

        node.text = text



# 주소 데이터가 없는 셀을 비우기 (서식/구조 유지)
def _clear_cell(cell):
    for node in cell.xpath('.//*[local-name()="t"]'):
        node.text = ""



# 제{{차수}}차를 제N차로 변경
def _fill_round_title(row, round_no):

    for node in row.xpath('.//*[local-name()="t"]'):
        if node.text:
            node.text = node.text.replace("{{차수}}",str(round_no))


# cellAddr를 다시 정리
def _update_cell_addresses(table):

    rows = table.xpath('./*[local-name()="tr"]')

    for row_index, row in enumerate(rows):

        cells = row.xpath('./*[local-name()="tc"]')

        for col_index, cell in enumerate(cells):

            cell_addr = cell.xpath('./*[local-name()="cellAddr"]')

            if not cell_addr:
                continue

            cell_addr[0].set("rowAddr",str(row_index))
            cell_addr[0].set("colAddr",str(col_index))




# =========================================
# Excel에서 만든 주소 데이터를 이용해 HWPX 생성
# =========================================
def create_address_hwpx(template_path, output_path, data):

    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

    if isinstance(data, str):
        data = json.loads(data)

    if not isinstance(data, list):
        raise ValueError("data는 JSON 배열(list)이어야 합니다.")

    if len(data) == 0:
        raise ValueError("입력 데이터가 없습니다.")


    groups = _group_by_round(data)

    with zipfile.ZipFile(template_path, "r") as zin:
        section_xml = zin.read("Contents/section0.xml")

        # xml 분석 parser (내부 줄바꿈, 공백 유지)
        parser = etree.XMLParser(remove_blank_text=False)

        # parser로 xml 객체로 변환 후 최상위 태그 root 가져오기 
        root = etree.fromstring(section_xml, parser)

        # 템플릿의 첫 번째 표를 차수/주소록 표의 기본 틀로 사용
        tables = root.xpath('.//*[local-name()="tbl"]')

        if not tables:
            raise ValueError("section0.xml에서 표(tbl)를 찾지 못했습니다.")


        template_table = tables[0]
        template_rows = template_table.xpath('./*[local-name()="tr"]')

        if len(template_rows) < 2:
            raise ValueError("첫 번째 표에 제목 행과 주소 행이 최소 2개 필요합니다.")


        # 템플릿 복사
        title_row_template = deepcopy(template_rows[0])
        address_row_template = deepcopy(template_rows[1])

        address_cells = address_row_template.xpath('./*[local-name()="tc"]')

        if len(address_cells) != 2:
            raise ValueError("주소 행의 셀 수가 2개가 아닙니다.")


        # 기존 행 제거
        for row in list(template_table.xpath('./*[local-name()="tr"]')):
            template_table.remove(row)


        # 차수별로 동적으로 행 생성
        for round_no, items in groups.items():

            # 1. 차수 제목 행
            title_row = deepcopy(title_row_template)
            _fill_round_title(title_row, round_no)
            template_table.append(title_row)

            # 2. 주소 카드 행
            for i in range(0, len(items), 2):

                left_item = items[i]

                right_item = None
                if i + 1 < len(items):
                    right_item = items[i + 1]

                new_row = deepcopy(address_row_template)
                cells = new_row.xpath('./*[local-name()="tc"]')

                # 왼쪽
                _replace_in_cell(cells[0], left_item)

                # 오른쪽
                if right_item is not None:
                    _replace_in_cell(cells[1], right_item)

                else:
                    _clear_cell(cells[1])

                template_table.append(new_row)


        # 표의 실제 행수를 HWPX 메타데이터에 반영
        actual_rows = template_table.xpath('./*[local-name()="tr"]')
        
        # rowCnt 갱신
        template_table.set("rowCnt", str(len(actual_rows)))

        # ★ cellAddr의 rowAddr / colAddr 갱신
        _update_cell_addresses(template_table)


        new_section_xml = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True
        )

        # HWPX ZIP 재생성
        with zipfile.ZipFile(output_path,"w",compression=zipfile.ZIP_DEFLATED) as zout:

            for info in zin.infolist():

                if info.filename == "Contents/section0.xml":
                    zout.writestr(info,new_section_xml)

                else:
                    zout.writestr(info, zin.read(info.filename))

    return str(output_path)



if __name__ == "__main__":

    template = r"C:\Users\dhkim\OneDrive\문서\UiPath\02_CreateLabel\Data\Temp\주소록_양식.hwpx"
    output = r"C:\Users\dhkim\OneDrive\문서\UiPath\02_CreateLabel\Data\Output\주소록_결과.hwpx"

    test_data = [
        {
            "차수": "1",
            "주소": "전북특별자치도 군산시 미장13길 33 107동 2001호",
            "보호자명": "신민진",
            "학생명": "장태현",
            "우편번호": "54096"
        },
        {
            "차수": "1",
            "주소": "군산시 궁포2로 100 101동 1001호",
            "보호자명": "김영희",
            "학생명": "홍길동",
            "우편번호": "54097"
        },
        {
            "차수": "1",
            "주소": "군산시 미장로 20",
            "보호자명": "이영희",
            "학생명": "김철수",
            "우편번호": "54098"
        },
        {
            "차수": "2",
            "주소": "군산시 수송동 123",
            "보호자명": "박영희",
            "학생명": "최민수",
            "우편번호": "54099"
        },
        {
            "차수": "2",
            "주소": "군산시 나운동 456",
            "보호자명": "최영희",
            "학생명": "이민수",
            "우편번호": "54100"
        }
    ]

    create_address_hwpx(
        template,
        output,
        test_data
    )

    print(f"★완료: {output}")
