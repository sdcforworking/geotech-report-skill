"""폴더 정리 이동 내역 엑셀에 한 줄 추가 (원래 위치 / 새 위치 / 비고).
기록 파일은 --log 또는 환경변수 MOVE_LOG로 지정. 엑셀이 열려 있으면(~$ 잠금 파일) 멈춤.
기록 시트 이름은 '이동 내역'(없으면 첫 시트).

사용:
  python log_move.py --log "D:\\과업\\보관\\00_폴더 정리 이동 내역.xlsx" 원래위치 새위치 "비고"
  (경로 구분은 / 로 넣어도 됨, 같은 경로 두 번 = 파일을 그 자리에서 고친 기록)"""
import argparse, os, sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
ap = argparse.ArgumentParser()
ap.add_argument('--log', default=os.environ.get('MOVE_LOG'))
ap.add_argument('src'); ap.add_argument('dst'); ap.add_argument('note')
a = ap.parse_args()
if not a.log:
    sys.exit('기록 파일을 --log 또는 MOVE_LOG로 지정하세요')
lock = os.path.join(os.path.dirname(a.log), '~$' + os.path.basename(a.log))
if os.path.exists(lock):
    sys.exit('기록 엑셀이 열려 있음 — 닫고 다시 실행')
wb = openpyxl.load_workbook(a.log)
ws = wb['이동 내역'] if '이동 내역' in wb.sheetnames else wb.worksheets[0]
ws.append([a.src.replace('/', '\\'), a.dst.replace('/', '\\'), a.note])
wb.save(a.log)
print('기록', ws.max_row, '|', a.src, '→', a.dst)
