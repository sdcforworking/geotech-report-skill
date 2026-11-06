"""한글(HWP) COM 자동화 시작 도우미 — '접근 허용' 창이 매번 뜨지 않게 보안 모듈을 먼저 등록함.

준비(한 번만): 한글 보안 모듈 DLL(예: C:\\HWPMODULE\\FilePathCheckerModuleExample.dll)을 받아 두고
레지스트리 HKCU\\Software\\HNC\\HwpAutomation\\Modules 에 문자열 값 FilePathCheckerModule = DLL 경로 를 등록.

사용 예:
    from hwp_com import open_hwp
    hwp = open_hwp(r'D:\\과업\\본문.hwpx', visible=False)
    ... hwp.HAction / hwp.GetText 등 ...
    hwp.Quit()
주의: 사용자가 같은 파일을 한글로 열어 두었으면 먼저 닫게 하고, 덮어쓰기 전에는 백업할 것."""
import win32com.client as win32


def open_hwp(path=None, visible=False):
    hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
    hwp.XHwpWindows.Item(0).Visible = visible
    if path:
        if not hwp.Open(path, '', 'forceopen:true'):
            raise RuntimeError(f'못 엶: {path}')
    return hwp


if __name__ == '__main__':
    import sys
    h = open_hwp(sys.argv[1] if len(sys.argv) > 1 else None)
    print('쪽 수', h.PageCount)
    h.Quit()
