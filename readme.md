# PyDocx Editor

Python으로 만든 Word(.docx) 파일용 GUI 데스크톱 편집기입니다. Tkinter(GUI) + python-docx(파일 입출력)로 동작하며, 별도의 무거운 프레임워크 없이 python app.py 한 줄로 실행할 수 있습니다.

## 주요 기능
문서 열기/저장: 기존 .docx 파일을 열어 편집하고 다시 저장 (새 문서 생성도 가능)
텍스트 서식: 굵게(Ctrl+B) · 기울임(Ctrl+I) · 밑줄(Ctrl+U) · 글꼴 · 크기 · 글자색
문단 스타일: Normal / Title / Heading 1~4 (문서 구조에 그대로 반영)
정렬: 왼쪽 / 가운데 / 오른쪽 / 양쪽 정렬
이미지 삽입: 원하는 위치에 이미지를 넣고, 저장 시 원본 해상도로 문서에 포함
표 삽입/편집: 행/열 지정 후 표를 만들고, 문서 내 표시된 위치를 더블클릭하면 셀 내용 편집 가능
찾기/바꾸기: Ctrl+F로 검색 및 일괄 치환
실행취소/다시실행: Ctrl+Z / Ctrl+Y
상태 표시줄: 단어 수, 글자 수, 커서 위치, 저장 여부(*) 표시

##설치

Python 3.9 이상 권장.

```bash
```pip install python-docx pillow

Tkinter는 파이썬 표준 라이브러리이지만 일부 Linux 배포판은 별도 설치가 필요합니다.

```bash
# Ubuntu/Debian
```sudo apt-get install python3-tk

# macOS (python.org 설치본이면 기본 포함, Homebrew라면)
```brew install python-tk

# Windows: 공식 설치 프로그램에 기본 포함되어 있어 별도 작업 불필요
실행
```bash
```py app.py

##사용 팁
문단 스타일과 정렬은 커서가 있는 줄(문단) 전체에 적용됩니다.
글꼴/크기/색상/굵게/기울임/밑줄은 드래그로 선택한 텍스트에 적용됩니다.
표는 문서 안에서 ⟦ TABLE:... ⟧ 형태의 자리표시자로 보이며, 이 줄을 더블클릭하면 표 편집 창이 열립니다. 저장 시 실제 Word 표로 변환됩니다.
기존 문서를 열면 굵게/기울임/글자색/문단 스타일/정렬/이미지/표가 최대한 그대로 복원됩니다. 다만 아주 복잡한 서식(다단 레이아웃, 각주, 목차 필드 등)은 단순화되어 표시될 수 있습니다.

##제한
표 안에 이미지나 별도 서식이 있는 셀은 텍스트만 보존됩니다.
한 문단 안에서 줄바꿈(soft line break)은 별도 문단으로 취급됩니다.
각주/미주, 목차(TOC) 필드, 머리글/바닥글은 편집 대상에 포함되지 않습니다.

## 저작권 및 오픈소스 라이선스 고지

이 프로그램은 다음과 같은 오픈소스 소프트웨어를 사용합니다.

- **Python**
  Copyright © 2001-2026 Python Software Foundation. All Rights Reserved.
  Licensed under the PSF License Agreement.
  https://docs.python.org/3/license.html

- **Tkinter / Tcl/Tk**
  Python 표준 라이브러리에 포함된 GUI 툴킷으로, Tcl/Tk를 기반으로 합니다.
  Copyright © Regents of the University of California, Sun Microsystems, Scriptics Corporation, ActiveState 외.
  Licensed under the Tcl/Tk License (BSD-style).
  https://www.tcl.tk/software/tcltk/license.html

- **python-docx**
  Copyright © 2013 Steve Canny, https://github.com/python-openxml/python-docx
  Licensed under the MIT License.

- **Pillow**
  Copyright © 2010 by Jeffrey A. Clark and contributors (Pillow), and
  Copyright © 1997-2011 by Secret Labs AB and Fredrik Lundh (원본 PIL).
  Licensed under the HPND License (MIT-CMU 계열, permissive license).
  https://github.com/python-pillow/Pillow/blob/main/LICENSE

각 라이브러리의 저작권 및 라이선스 조건은 해당 프로젝트의 원본 라이선스 파일을 따릅니다. 이 프로그램 자체의 소스 코드에 대한 별도 저작권 표기는 상단 라이선스 섹션을 참고하세요.

연락처
이메일: lotus031315@gmail.com
포트폴리오: https://joon041415.github.io/
