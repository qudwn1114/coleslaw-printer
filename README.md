## 프로젝트 소개

 ESC/POS 프린트
 
## 기술 스택

|Category| - |
| --- | --- |
|Language|Python v3.11.11|
|Framework|Flask v3.1.0|



### 사용법

Execute the following lines to properly clone and run the project.  

COMMAND
- $ conda create -n ["env"] python=3.11.11
- $ conda activate ["env"]
- $ pip install -r requirements.txt
- $ python esc_pos.py

RELEASE 
- `.exe` 파일을 **릴리즈** 목록에서 다운로드 하여 실행 하 실 수 있습니다.


실행파일 만들기
- $ pyinstaller --onefile --windowed --hidden-import plyer.platforms.win.notification esc_pos.py
