import pyautogui
import pyautogui as pg
import pyperclip
from time import sleep

x, y = pg.position()
print('mouse_position:',x,y)

# login - id 입력 720 435

# 마우스 특정 위치로 이동
pg.moveTo(825, 670)

# 마우스 클릭
pg.click()

# keyboard 입력
pg.typewrite('h01041303675@gmail.com', interval=0.1)

# password 입력
# 859 700
pg.moveTo(859,700)
pg.click()
pg.typewrite('jy814092@@', interval=0.1)

# 로그인 버튼
pg.moveTo(925, 749)
sleep(0.5)
pg.click()


# 복사 봍여넣기
# pyperclip.copy('마케팅 메세지 00고객님 오늘의 성과는~~')
# pg.hotkey('ctrl', 'v')
