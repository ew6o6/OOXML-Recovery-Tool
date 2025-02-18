import re
import binascii
import json
import os
import zlib
from bs4 import BeautifulSoup
from tabulate import tabulate
import csv
from docx import Document
import warnings # BeautifulSoup 경고 메시지 무시

# 4) 손상된 local file의 file data(hex형태) 최대한 압축 해제
def decompress_deflate_hex(hex_string):
    # 문자열(hex 값)을 바이트열로 변환
    compressed_data = bytes.fromhex(hex_string)
    #decompressor
    decompressor = zlib.decompressobj(-zlib.MAX_WBITS)

    try:
        decompressed_data = decompressor.decompress(compressed_data)
        # decode decompressed file data 
        decompressed_string = decompressed_data.decode("utf-8")
    except UnicodeDecodeError:
        # 손상된 compressed_data 값을 뒤에서부터 1byte씩 삭제하며 계속 decompress 시도
        for i in range(1, len(compressed_data)):
            try:
                decompressed_data = zlib.decompress(compressed_data[:-i], -zlib.MAX_WBITS)
                # decode decompressed file data
                decompressed_string = decompressed_data.decode("utf-8")
                break
            except UnicodeDecodeError:
                continue
    
    return decompressed_string

# 3) local header에서 extra field 값 제거(i.g. 파일명만 남김)
def extract_file_name(decoded_local_file_name):
    #텍스트가 아닌 문자열() 정규 표현식 지정 후 뒤에서부터 탐색
    pattern_for_non_text = re.search(r'[\x00-\x1F\x7F]', decoded_local_file_name[::-1])
    if pattern_for_non_text:
        # 일치하는 위치까지의 문자열만 남김
        decoded_local_file_name = decoded_local_file_name[-pattern_for_non_text.end()+1:]
    return decoded_local_file_name

# 2) xml_name_with_hex의 local_file_name index 값(localfile_name)을 utf-8로 디코딩
def decode_utf8(local_file_name_hex):
    #hex_str을 16진수로 변환
    local_file_name_hex = binascii.unhexlify(local_file_name_hex)
    #hex_str을 utf-8로 디코딩
    decoded_local_file_name = local_file_name_hex.decode('utf-8', 'ignore')
    return decoded_local_file_name

# 6) local_file_data index값(file data) 
def decode_local_file_data(json_file_name):        
    for item in json_file_name:
        if item['local_file_name'].startswith('word/media/') or item['local_file_name'].startswith('xl/media/'):
            continue

        hex_string = item['local_file_data']
        if hex_string:
            try:
                decompressed_string = decompress_deflate_hex(hex_string)
                item['local_file_data'] = decompressed_string
                #print(item['local_file_name'], item['local_file_data'])
            except zlib.error:
                # 손상된 local_file_data 값을 뒤에서부터 2자리씩 삭제하며 decompress 시도
                for i in range(2, len(hex_string), 2):
                    try:
                        decompressed_string = decompress_deflate_hex(hex_string[:-i])
                        item['local_file_data'] = decompressed_string
                        break
                    except zlib.error:
                        continue
            except UnicodeDecodeError:
                # 손상된 local_file_data 값을 뒤에서부터 2자리씩 삭제하며 decode 시도
                for i in range(2, len(hex_string), 2):
                    try:
                        decoded_str = decode_utf8(hex_string[:-i])
                        item['local_file_data'] = decoded_str
                        break
                    except UnicodeDecodeError:
                        continue

#xml 텍스트만 추출(for docx)
def extract_xml_text_for_docx(with_xml_tag_data: str) -> str:
    #old_version
    # #xml tag 제거를 위해 '<'로 시작하고 '>'로 끝나는 문자열 정규표현식 패턴 정의
    # pattern =  r'<[^>]+>'
    
    # #re 모듈의 sub() 함수를 이용하여 xml 태그를 빈 문자열로 치환
    # clean_text = re.sub(pattern, "", with_xml_tag_data)
    
    # #xml tag 제거된 문자열을 BeautifulSoup 모듈의 get_text() 함수를 이용하여 텍스트 추출
    # xml_text = BeautifulSoup(clean_text, "lxml").get_text()
    
    #new_version_v1(띄어쓰기 및 형식 정보 활용)
    # # xml tag 제거를 위해 '<'로 시작하고 '>'로 끝나는 문자열 정규표현식 패턴 정의
    # pattern = r'<[^>]+>'
    
    # # re 모듈의 sub() 함수를 이용하여 xml 태그를 빈 문자열로 치환
    # clean_text = re.sub(pattern, '', with_xml_tag_data)
    
    # # xml tag 제거된 문자열을 BeautifulSoup 모듈의 get_text() 함수를 이용하여 텍스트 추출
    # soup = BeautifulSoup(clean_text, 'html.parser')
    # xml_text = soup.get_text()

    #new_version_v2
    #  # xml tag 제거를 위해 '<'로 시작하고 '>'로 끝나는 문자열 정규표현식 패턴 정의
    # pattern = r'<[^>]+>'

    # # re 모듈의 sub() 함수를 이용하여 xml 태그를 빈 문자열로 치환
    # clean_text = re.sub(pattern, '', with_xml_tag_data)

    # # 개행을 스페이스로 대체하여 BeautifulSoup 모듈의 get_text() 함수를 이용하여 텍스트 추출
    # xml_text = BeautifulSoup(clean_text, 'html.parser').get_text(separator=' ')

    # # BeautifulSoup을 사용하여 XML 파싱(띄어쓰기만 활용)
    # soup = BeautifulSoup(with_xml_tag_data, "xml")

    # # 텍스트 추출
    # BeautifulSoup의 경고 메시지를 무시
    warnings.filterwarnings("ignore", category=UserWarning, module="bs4")
    # xml_text = soup.get_text(separator=' ')
    soup = BeautifulSoup(with_xml_tag_data, "lxml")
    
    # 개행 문자 처리: '<w:p>' 태그를 기준으로 문단을 구분하고 개행 문자를 띄어쓰기로 대체
    for paragraph in soup.find_all('w:p'):
        paragraph.string = ' '.join(paragraph.stripped_strings)
    
    # 텍스트 추출
    xml_text = soup.get_text(separator='\n')  # 문단 간 개행 문자로 구분
    return xml_text

#img 파일 추출
def extract_img_file(local_file, output_path):
    for item in local_file:
        if 'img_ext' in item and item['img_ext'] is not None:
            img_file_name = os.path.basename(item['local_file_name'])
            img_path = os.path.join(output_path, img_file_name)

            with open(img_path, 'wb') as f:
                f.write(binascii.unhexlify(item['local_file_data']))

#metadata 추출
def extract_metadata(local_file, output_directory):
    tags = {
    'dc:creator': 'creator',
    'cp:lastModifiedBy': 'lastModifiedBy',
    'cp:revision': 'revision',
    'dcterms:created': 'created',
    'dcterms:modified': 'modified'
    }

    results = []
    results.append("[문서 metadata]")

    for item in local_file:
        if item['local_file_name'] == 'docProps/core.xml':
            all_tags_missing = True
            for tag, name in tags.items():
                pattern = f"<{tag}.*?>(.*?)</{tag}>"
                match = re.search(pattern, item['local_file_data'])
                if match and match.group(1):
                    all_tags_missing = False
                    value = match.group(1)
                    results.append(f"{name} : {value}(<{tag}>{value}</{tag}>)")

            if all_tags_missing:
                results = ["손상 정도가 심각하여 metadata가 존재하지 않습니다."]

    metadata_output_path = os.path.join(output_directory, 'metadata', 'metadata.txt')
    metadata_output_dir = os.path.dirname(metadata_output_path)

    os.makedirs(metadata_output_dir, exist_ok=True)

    metadata_output_text = '\n'.join(results)

    with open(metadata_output_path, 'w', encoding='utf-8') as f:
        f.write(metadata_output_text)
    
# 파일 읽은 후 pk를 기준으로 local file 분할 -> local file header의 .xml/.rels 기준으로 localfile_name과 file data 분할 -> json 형식으로 저장
def get_file_hex(file_path):
    base_directory = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)
    file_base_name, _ = os.path.splitext(file_name)
    #input file의 확장자는 아래의 파일 식별 단계에서 판별
    global file_ext
    file_ext = ""
    #file 읽기
    with open(file_path, 'rb') as f:
        data = f.read()
            
    # pk signature(\\x50\\x4B\\x03\\x04) 탐지 -> 시작 위치 리스트에 저장
    pattern_for_pk_sig = re.compile(b'\\x50\\x4B\\x03\\x04')
    matches = pattern_for_pk_sig.finditer(data)

    if not matches:
        print("The input file is not a valid OOXML-based MS document. Data extraction has been terminated.")
        return None  # 데이터 추출 진행하지 않음
    
    # 입력 파일이 OOXML 파일인가?(OOXML 파일의 pk signature 이후에 오는 압축 version 정보가 0x14 0x00인가?) -> 아니면 함수 종료
    pattern_for_version = re.compile(b'\\x14\\x00')
    first_match = next(matches, None)
    if not (first_match and pattern_for_version.match(data, first_match.end())):  # OOXML 파일이 아닌 경우
        print("The input file is not a valid OOXML-based MS document. Data extraction has been terminated.")
        return None  # 함수 종료


# [Content_Types].xml을 포함하여 아래의 local file name list에 존재하는 local file 중 하나라도 존재하는지 확인
    local_file_names = [
        b'[Content_Types].xml',
        b'_rels/.rels',
        b'docProps/app.xml',
        b'docProps/core.xml',
        b'styles.xml',
        b'word/document.xml',
        b'xl/sharedStrings.xml',
        re.compile(b'xl/worksheets/sheet\d+\.xml'),
        b'/media/'
    ]

    found_local_file = any(local_file_name in data for local_file_name in local_file_names)

    if not found_local_file:
        print("The input file is not a valid OOXML-based MS document. Data extraction has been terminated.")
        return None

    # word/document.xml 파일이 존재하면 Word(docx) 파일로 판별
    pattern_for_word_document = re.compile(b'word/document.xml')
    word_document_match = pattern_for_word_document.search(data)

    # xl/sharedStrings.xml 또는 xl/worksheets/sheet1.xml 파일이 존재하면 Excel(xlsx) 파일로 판별
    pattern_for_shared_strings = re.compile(b'xl/sharedStrings.xml')
    shared_strings_match = pattern_for_shared_strings.search(data)
    pattern_for_worksheet = re.compile(b'xl/worksheets/sheet\d+\.xml')
    worksheet_match = pattern_for_worksheet.search(data)

    file_ext = '.docx' if word_document_match else '.xlsx'

    # 만약 Word(docx) 또는 Excel(xlsx) 파일이지만, 본문 내용이 없는 경우 아래의 문구 출력 후 종료
    if not (word_document_match or (shared_strings_match and worksheet_match)):
        print(f"The input file is recognized as an {file_ext} file, but there is no recoverable document content.")
        return None
    # 리스트에 시작 위치 저장
    positions = [match.start() for match in matches]

    # hex 분할 : data의 시작부터 첫 번째 일치하는 위치 부분 문자열 제외
    parts = []
    for i, j in zip([0]+positions, positions+[None]):
        part = data[i:j]
        #part가 빈 문자열이 아닌 경우 == True
        if part:
            parts.append(part)

    # parts에서 .xml 또는 .rels 탐지 -> .xml/.rels별 hex 추출(json 형식)
    pattern_for_xml = re.compile(b'\\x2E\\x78\\x6D\\x6C')
    pattern_for_rels = re.compile(b'\\x2E\\x72\\x65\\x6C\\x73')
    patterns_for_img = {
    'emf': re.compile(b'\\x2E\\x65\\x6D\\x66'),
    'wmf': re.compile(b'\\x2E\\x77\\x6D\\x66'),
    'jpg': re.compile(b'\\x2E\\x6A\\x70\\x67'),
    'jpeg': re.compile(b'\\x2E\\x6A\\x70\\x65\\x67'),
    'jfif': re.compile(b'\\x2E\\x6A\\x66\\x69\\x66'),
    'jpe': re.compile(b'\\x2E\\x6A\\x70\\x65'),
    'png': re.compile(b'\\x2E\\x70\\x6E\\x67'),
    'bmp': re.compile(b'\\x2E\\x62\\x6D\\x70'),
    'dib': re.compile(b'\\x2E\\x64\\x69\\x62'),
    'rle': re.compile(b'\\x2E\\x72\\x6C\\x65'),
    'gif': re.compile(b'\\x2E\\x67\\x69\\x66'),
    'emz': re.compile(b'\\x2E\\x65\\x6D\\x7A'),
    'wmz': re.compile(b'\\x2E\\x77\\x6D\\x7A'),
    'tif': re.compile(b'\\x2E\\x74\\x69\\x66'),
    'tiff': re.compile(b'\\x2E\\x74\\x69\\x66\\x66'),
    'svg': re.compile(b'\\x2E\\x73\\x76\\x67'),
    'ico': re.compile(b'\\x2E\\x69\\x63\\x6F'),
    'heif': re.compile(b'\\x2E\\x68\\x65\\x69\\x66'),
    'heic': re.compile(b'\\x2E\\x68\\x65\\x69\\x63'),
    'hif': re.compile(b'\\x2E\\x68\\x69\\x66'),
    'avif': re.compile(b'\\x2E\\x61\\x76\\x69\\x66'),
    'webp': re.compile(b'\\x2E\\x77\\x65\\x62\\x70')
}

    local_file_in_hex = []

    for part in parts:
        # 패턴(확장자) 검사 수행
        match_xml = pattern_for_xml.search(part)
        match_rels = pattern_for_rels.search(part)

        match_img = None
        img_ext = None
        for key, pattern in patterns_for_img.items():
            temp_match = pattern.search(part)
            if temp_match:
                # 이미지에 대한 일치 위치 업데이트
                if not match_img or temp_match.end() > match_img.end():
                    match_img = temp_match
                    img_ext = key

        
        # 패턴 검사 -> 가장 뒤에 있는 위치를 기준으로 문자열 분할
        if match_xml or match_rels or match_img:
            if match_xml and match_rels and match_img:
                
                index = max(match_xml.end(), match_rels.end(), match_img.end())
            elif match_xml and match_rels:
                index = max(match_xml.end(), match_rels.end())
            elif match_xml and match_img:
                index = max(match_xml.end(), match_img.end())
            elif match_rels and match_img:
                index = max(match_rels.end(), match_img.end())
            elif match_xml:
                index = match_xml.end()
            elif match_rels:
                index = match_rels.end()
            else:
                index = match_img.end()
            

            # 일치하는 위치를 기준으로 문자열 분할
            local_file_name = part[:index]
            local_file_data = part[index:]

            # extra_field 시작 위치 및 크기 offset
            extra_field_start_offset = index + 1
            
            # extra_field 시작 위치 및 크기 계산
            extra_field_length = int.from_bytes(part[28:30], 'little')

            # extra_field가 있는 경우, 그 값을 제외하고 file data 저장
            if extra_field_length > 0:
                local_file_data = part[extra_field_start_offset + extra_field_length -1:]
            else:
                local_file_data = part[extra_field_start_offset-1:]

            # 분할된 문자열을 json 형식으로 저장
            if match_xml or match_rels:
                local_file_in_hex.append({
                    'local_file_name': extract_file_name(decode_utf8(binascii.hexlify(local_file_name).decode())),
                    'local_file_data': binascii.hexlify(local_file_data).decode()
                })
            else:
                local_file_in_hex.append({
                    'local_file_name': extract_file_name(decode_utf8(binascii.hexlify(local_file_name).decode())),
                    'local_file_data': binascii.hexlify(local_file_data).decode(),
                    'img_ext': img_ext
                })

        else:
            # 일치하는 위치가 없으면 전체 문자열만 저장
            local_file_in_hex.append({
                'local_file_name': extract_file_name(decode_utf8(binascii.hexlify(part).decode())),
                'local_file_data': '',
                'img_ext': None
            })

    decode_local_file_data(local_file_in_hex)

    # Output directories and file paths
    output_directory = os.path.join(base_directory, f'output_{file_base_name}')
    #file_name에서 확장자 제거
    file_name_without_ext = file_name.rsplit('.', 1)[0]
    json_output_path = os.path.join(output_directory, f'{file_name_without_ext}_localFiles.json')
    img_output_dir = os.path.join(output_directory, 'img')
    #metadata_output_dir = os.path.join(output_directory, 'metadata')

    os.makedirs(output_directory, exist_ok=True)
    os.makedirs(img_output_dir, exist_ok=True)

    # JSON 저장
    i = 1
    while os.path.exists(json_output_path):
        json_output_path = os.path.join(output_directory, f'{file_name}({i}).json')
        i += 1

    with open(json_output_path, 'w', encoding='UTF-8-sig') as f:
        f.write(json.dumps(local_file_in_hex, ensure_ascii=False, indent="\t"))

    extract_img_file(local_file_in_hex, img_output_dir)
    extract_metadata(local_file_in_hex, output_directory)

    # 파일 확장자를 판별하여 file_ext 변수 설정
    if 'word/document.xml' in [item['local_file_name'] for item in local_file_in_hex]:
        file_ext = '.docx'
    elif 'xl/sharedStrings.xml' in [item['local_file_name'] for item in local_file_in_hex]:
        file_ext = '.xlsx'
    else:
        file_ext = None

    if file_ext == '.docx':
        print("actual file extension : docx")
        process_extracted_data(local_file_in_hex, file_path)
    elif file_ext == '.xlsx':
        print("actual file extension : xlsx")
        process_extracted_data(local_file_in_hex, file_path)
    else:
        print("It is not an MS Office file.")

    return local_file_in_hex

def extract_data_from_shared_strings(shared_strings_xml):
    pattern = r'<t>([^<]+)</t>'
    return re.findall(pattern, shared_strings_xml)

def extract_data_from_document(document_xml):
    pattern = r'<w:t>([^<]+)</w:t>'
    return re.findall(pattern, document_xml)

def extract_data_from_sheet(sheet_xml, shared_strings):
    pattern = r'<c r="(\w+)"(?: t="(\w+)")?>(?:<v>([^<]*)</v>)?</c>'
    matches = re.findall(pattern, sheet_xml)

    mapped_data = {}
    unmapped_data = set(shared_strings)  # 초기 상태에서 모든 shared_strings은 매핑되지 않은 것으로 간주

    for cell_ref, data_type, value in matches:
        row = int(re.search(r'\d+', cell_ref).group())
        col = re.search(r'[A-Z]+', cell_ref).group()

        if data_type == 's':
            if int(value) < len(shared_strings):
                mapped_data.setdefault(row, {})[col] = shared_strings[int(value)]
                unmapped_data.discard(shared_strings[int(value)])  # 매핑된 데이터는 제거
            else:
                unmapped_data.add("ERROR")
        else:
            mapped_data.setdefault(row, {})[col] = value

    return mapped_data, unmapped_data


# CSV 파일로 저장
def save_unmapped_data_to_csv(unmapped_data, output_directory, filename="unmapped_data.csv"):
    output_file_path = os.path.join(output_directory, filename)
    with open(output_file_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(["Unmapped Data from sharedStrings.xml"])
        for data in unmapped_data:
            csv_writer.writerow([data])

def process_extracted_xlsx_data(local_file_xml, file_path):
    sharedStrings_xml = ""
    sheet_xmls = []
    all_mapped_values = []

    # 모든 XML 파일에서 sharedStrings.xml 및 sheet[N].xml 탐색
    for item in local_file_xml:
        if item['local_file_name'] == 'xl/sharedStrings.xml':
            sharedStrings_xml = item['local_file_data']
        elif re.match(r'xl/worksheets/sheet\d+.xml', item['local_file_name']):
            sheet_xmls.append(item['local_file_data'])

    # sharedStrings.xml 및 sheet[N].xml 처리
    shared_strings = extract_data_from_shared_strings(sharedStrings_xml)
    output_directory = os.path.join(os.path.dirname(file_path), f'output_{os.path.basename(file_path).rsplit(".", 1)[0]}')
    os.makedirs(output_directory, exist_ok=True)

    for i, sheet_xml in enumerate(sheet_xmls, start=1):
        # print(f'\nProcessing sheet {i}:')
        mapped_data, unmapped_data = extract_data_from_sheet(sheet_xml, shared_strings)
        for cell_data in mapped_data.values():
            all_mapped_values.extend(cell_data.values())

        # Excel 표 출력 및 CSV 파일 생성
        table_filename = f"output_damaged_sheet{i}.csv"
        display_and_save_table_to_csv(mapped_data, table_filename, unmapped_data, file_path)
    print(f"\n[output path]\n{output_directory}")

    # 매핑되지 않은 데이터 찾기
    all_mapped_values_set = set(all_mapped_values)
    unmapped_data = [s for s in shared_strings if s not in all_mapped_values_set]
    if unmapped_data:
        save_unmapped_data_to_csv(unmapped_data, output_directory, "unmapped_data.csv")
    #print(f"\n[output path]\n{output_directory}")

# [Excel] table화하여 CSV 파일로 저장
def display_and_save_table_to_csv(mapped_data, filename, unmapped_data=None, file_path=None):
    if not mapped_data:
        return

    rows = sorted(mapped_data.keys())
    columns = sorted({col for row_data in mapped_data.values() for col in row_data.keys()})

    table = []
    headers = [""] + columns
    for row in rows:
        table_row = [row]
        for col in columns:
            table_row.append(mapped_data.get(row, {}).get(col, ""))
        table.append(table_row)

    table_str = tabulate(table, headers=headers, tablefmt='grid')

    if file_path:
        output_directory = os.path.join(os.path.dirname(file_path), f'output_{os.path.basename(file_path).rsplit(".", 1)[0]}')
        os.makedirs(output_directory, exist_ok=True)
        filename = os.path.join(output_directory, filename)
        #print(f"\n[output path]\n{output_directory}")
        # CSV 파일로 저장
        with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            
            for row_data in table:
                writer.writerow(row_data)


#[Word] docx 파일로 저장
def save_text_to_docx(text, filename="output.docx"):
    
    docx_document = Document()
    docx_document.add_paragraph(text)
    #print(f"\n[output file]\n{filename}")
    docx_document.save(filename)

#localfile_name이 'word/document.xml'인 local_file_data에서 텍스트 추출
def process_extracted_docx_data(local_file_xml, file_path):
    
    # xml_name_with_hex가 None이 아닌 경우(==OOXML 기반 MS 파일인 경우)에만 해당 함수를 호출
    if not local_file_xml:
        return
    
    base_directory = os.path.dirname(file_path)
    file_base_name = os.path.basename(file_path).rsplit('.', 1)[0]

    
    # 아래와 같이 output_directory 값을 설정합니다.
    output_directory = os.path.join(base_directory, f'output_{file_base_name}')
    os.makedirs(output_directory, exist_ok=True)

    for item in local_file_xml:
        if item['local_file_name'] == 'word/document.xml':
            local_file_name = item['local_file_name']
            #item['local_file_data']에 값이 없는 경우(CD이거나 압축해제된 데이터가 없는 local file)는 아래의 print문 실행X
            if item['local_file_data']:
                #xml tag 제거
                text = extract_xml_text_for_docx(item['local_file_data'])
                
                #docx 파일로 저장
                save_text_to_docx(text, filename=os.path.join(output_directory, "output_damaged.docx"))
            print(f"\n[output path]\n{output_directory}")

def process_extracted_data(local_file_xml, file_path):
    if not local_file_xml:
        return

    if file_ext == '.xlsx':
        process_extracted_xlsx_data(local_file_xml, file_path)
    elif file_ext == '.docx':
        process_extracted_docx_data(local_file_xml, file_path)
    else:
        print("지원되지 않는 파일 형식입니다.")

# 사용자로부터 경로 입력받아 디렉토리 또는 파일인지 확인 후 처리
def process_directory_or_file(path):
    if os.path.isdir(path):
        for filename in os.listdir(path):
            file_path = os.path.join(path, filename)
            if os.path.isfile(file_path):
                # 각 파일 처리
                process_file(file_path)
    elif os.path.isfile(path):
        # 파일 처리
        base_directory = os.path.dirname(path)
        process_file(path)
    else:
        print(f"제공된 경로 '{path}'은(는) 유효한 디렉토리 또는 파일이 아닙니다.")


# 사용자로부터 파일 경로 입력받기
def process_file(file_path):
    #file_path에서 파일명만 추출
    file_name = os.path.basename(file_path)
    print('------------------------------------')
    print(f'[input file name]\n{file_name}\n')
    local_file_in_hex = get_file_hex(file_path)
    print('------------------------------------\n')
    pass

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# 
input_path = input("Enter directory or file path : ")
process_directory_or_file(input_path)