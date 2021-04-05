# samsung_apply_fill_macro

삼성 채용 성적 입력 매크로  
테스트 환경: `Windows10 Chrome`  

정수 단위 학점 외 입력은 수동으로 하여야 합니다.  
리눅스에서 작동하지 않습니다.  

본 프로그램으로 발생하는 불이익은 저작권자가 책임지지 않습니다.  

# How to use

1. python3.6 설치. 예시링크(https://winpython.github.io/)

2. pip install -r requirements.txt

3. data.xlsx 형식으로 성적 엑셀 입력  
   학교란은 편입한 경우 학적에 뜨는 학교 이름 순서대로 1번 2번 ...  
   해당사항이 없는 경우는 그냥 1로 적으면 된다.  

4. 미리 삼성 채용 홈페이지-채용 지원서 작성-이수교과목 페이지 접속

5. run.py "엑셀 파일" 실행  
   ex) run.py data.xlsx  
   ex) run.py seoul.xlsx  

6. 프로그램 실행 후, config.yaml 파일의 setup_delay_time 안에 과목 입력란 중 과정 누르고 엔터  
   매크로 실행 중 중간에 중단되어 과목 수 카운트 수가 달라진다면 config.yaml 파일의 start_num 수정
![](docs/1.PNG)
![](docs/2.PNG)
7. 네트워크 지연속도나 안정성을 고려한다면 config.yaml 파일의 delay_time 조절  
   학교명 서버 데이터를 가져오는 것이 느리므로 안정적으로 실행하려면 run.py 내의 서버 딜레이를 1초로 설정