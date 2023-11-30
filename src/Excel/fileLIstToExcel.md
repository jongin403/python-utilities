gpt prompt

이름에 file 이 들어간 것은 파일이고,
이름에 folder 가 들어간 것은 폴더라고 판단해줘

폴더 구조가 1번과 같을 때

1번
/fileA
/folderA
/folderA/folderAA/fileAAA
/folderA/folderAA/fileAAB
/folderB/fileBB

2번과 같이 각 파일의 경로를 출력 가능하도록 변경해줘
| 는 column 을 구분하는 표시고 [] 는 빈 cell 을 의미해

2번
fileA
folderA | folderAA | fileAAA
[] | [] | fileAAB
folderB | fileBB

폴더의 경로를 기반으로 중복을 검사해서 중복된 폴더 경로일 경우 빈 cell 로 나타내줘
폴더 안에 파일과 폴더가 없을 경우에는 표시를 해줘야하고
폴더 안에 파일과 폴더가 있는 경우에는 하위 폴더를 탐색해야해
