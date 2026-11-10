# -*- coding: utf-8 -*-
"""파이썬이 시작될 때 자동으로 읽혀(settings.json env.PYTHONPATH) Office 문서 저장을 office_guard 판단에 태운다.

openpyxl·python-docx·python-pptx·xlsxwriter·dsdxlsx 는 모두 zipfile 로 쓴다 - zipfile 한 곳과 shutil.copyfile·os.replace 만 지키면 된다.
문서 라이브러리를 쓰지 않는 파이썬에는 시작 비용이 없다(해당 모듈이 불려올 때만 끼어든다).
의도한 거부(OverwriteRefused) 외의 내부 오류는 모두 무시한다 - 이 파일 때문에 빌더가 깨지면 안 된다.
"""
import os
import sys

_EXT = (".xlsx", ".xlsm", ".docx", ".pptx", ".dsd")
_og = None
_derived = {}   # 불러온 문서를 다른 경로에 저장했을 때: 저장 경로 -> 표식 (임시 저장 후 교체하는 부분 편집을 알아보려고)
_saved = {}     # 이 프로세스가 코드 소유로 저장한 파일 -> 해시 (같은 프로세스의 COM 후처리를 알아보려고)
_built = {}     # 이 프로세스가 저장한 문서 (빌더 이력용)
_made = {}      # 이 프로세스가 zip 으로 새로 쓴 문서 아닌 이름의 파일(.tmp 등) -> 표식 (임시로 쓰고 옮기는 빌더를 알아보려고)
_tainted = set()  # 사람 문서를 문서 아닌 이름으로 복사한 경로 (그 사본을 다시 문서로 옮겨도 코드 소유가 되지 않게)
_com_wait = {}  # COM 으로 저장했지만 열려 있어 해시를 못 읽은 파일 -> (경로, 판단)


def _key(p):
    return os.path.normcase(os.path.abspath(os.fsdecode(os.fspath(p))))


def _stamp(p):
    st = os.stat(p)
    return (st.st_mtime_ns, st.st_size)


def _guard():
    global _og
    if _og is None:
        import importlib.util
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        spec = importlib.util.spec_from_file_location("office_guard", os.path.join(here, "office_guard.py"))
        m = importlib.util.module_from_spec(spec)
        sys.modules.setdefault("office_guard", m)
        spec.loader.exec_module(m)
        import atexit
        atexit.register(_at_exit)
        _og = m
    return _og


def _target(p):
    try:
        s = os.fsdecode(os.fspath(p))
    except TypeError:
        return None
    name = os.path.basename(s).lower()
    return s if name.endswith(_EXT) and not name.startswith("~$") else None


def _before(path, source):
    g = _guard()
    if _com_wait:
        _flush_com()
    try:
        return g.before_write(path, source)
    except g.OverwriteRefused:
        raise
    except Exception:
        return None


def _after(path, v, source=None):
    if v is None:
        return
    try:
        cur = _guard().after_write(path, v)
        key = os.path.abspath(path)
        if not (_guard().EXEMPT_DIRS & {x.lower() for x in os.path.normpath(key).split(os.sep)}):
            _built[key] = True
        if v["own"] and cur:
            _saved[key] = cur
        else:
            _saved.pop(key, None)
        if source and not _guard().same_path(source[0], path):
            _derived[os.path.normcase(key)] = source
        if cur:
            _com_wait.pop(os.path.normcase(key), None)
    except Exception:
        pass


def _flush_com():
    for k, (p, v) in list(_com_wait.items()):
        _after(p, v)


def _src_owned(src):
    """복사·이동할 원천이 코드가 만든 내용인지. 모르면 False."""
    try:
        k = _key(src)
        if k in _made:
            return _made[k] == _stamp(k)
        if k in _tainted:
            return False
        if not _target(src):
            return True                      # 문서 아닌 이름의 파일은 출처를 알 수 없다 - 예전처럼 코드가 만든 것으로 본다
        return bool(_guard().owned_content(k))
    except Exception:
        return False


def _carry(v, src, moved=False):
    """복사·이동은 내용의 주인을 바꾸지 않는다. 사람 문서의 사본은 그 경로에서만 코드 소유, 옮긴 것은 계속 사람 파일."""
    if v is None or not v.get("own") or _src_owned(src):
        return v
    return dict(v, own=False if moved else "at")


def _plain_copy(src, dst):
    """문서를 문서 아닌 이름으로 복사했을 때 원천이 사람 문서면 표시해 둔다."""
    try:
        if _target(src) and not _src_owned(src):
            _tainted.add(_key(dst))
            _made.pop(_key(dst), None)
    except Exception:
        pass


def _at_exit():
    """같은 스크립트가 저장 뒤 COM 으로 다듬었으면 다듬은 내용도 코드 소유로 올린다."""
    try:
        _flush_com()
    except Exception:
        pass
    for p, sha in list(_saved.items()):
        try:
            cur = _og.sha1(p)
            if cur != sha:
                _og.remember(cur)
        except Exception:
            pass
    try:
        script = sys.argv[0] if sys.argv else ""
        if script.lower().endswith(".py") and _built:
            here = os.path.dirname(os.path.abspath(script))
            mods = [getattr(m, "__file__", None) for m in list(sys.modules.values())]
            local = [m for m in mods if m and m.lower().endswith(".py") and os.path.dirname(os.path.abspath(m)) == here]
            _og.record_build(script, list(_built), local)
    except Exception:
        pass


def _find_source():
    """저장 중인 문서 객체의 표식을 호출 스택에서 찾는다."""
    f = sys._getframe(2)
    for _ in range(15):
        if f is None:
            return None
        for k in ("self", "workbook"):
            src = getattr(f.f_locals.get(k), "_og_src", None)
            if src:
                return src
        f = f.f_back
    return None


def _tag(obj, path):
    t = _target(path)
    if t:
        try:
            obj._og_src = _guard().source_of(t)
        except Exception:
            pass


def _patch_zipfile(m):
    Z = m.ZipFile
    orig_init, orig_close = Z.__init__, Z.close

    def __init__(self, file, mode="r", *a, **k):
        t = _target(file) if mode in ("w", "x", "a") else None
        if not t:
            made = None
            if mode in ("w", "x", "a"):
                try:
                    made = _key(file)
                    if mode == "a" and _made.get(made) != _stamp(made):
                        made = None
                except Exception:
                    made = None
            orig_init(self, file, mode, *a, **k)
            if made:
                self._og_made = made
            return
        source = None
        if mode == "a" and os.path.isfile(t):
            source = _guard().source_of(t)
        elif mode != "a":
            source = _find_source()
        v = _before(t, source)
        orig_init(self, file, mode, *a, **k)
        self._og = (t, v, source)

    def close(self):
        og = self.__dict__.pop("_og", None)
        made = self.__dict__.pop("_og_made", None)
        orig_close(self)
        if og:
            _after(*og)
        if made:
            try:
                _made[made] = _stamp(made)
                _tainted.discard(made)
            except Exception:
                pass

    Z.__init__, Z.close = __init__, close


def _derived_source(src):
    try:
        return _derived.get(os.path.normcase(os.path.abspath(os.fsdecode(os.fspath(src)))))
    except TypeError:
        return None


def _patch_shutil(m):
    orig = m.copyfile

    def copyfile(src, dst, *a, **k):
        t = _target(dst)
        if not t:
            r = orig(src, dst, *a, **k)
            _plain_copy(src, dst)
            return r
        source = _derived_source(src)
        v = _carry(_before(t, source), src)
        r = orig(src, dst, *a, **k)
        _after(t, v)
        return r

    m.copyfile = copyfile

    # Windows 의 파이썬 3.13 copy2 는 copyfile 을 거치지 않는다. move 는 기본 인자로 원래 copy2 를 쥐고 있다.
    orig2 = m.copy2

    def _copy2(moved, src, dst, *a, **k):
        d = os.path.join(dst, os.path.basename(os.fsdecode(os.fspath(src)))) if os.path.isdir(dst) else dst
        t = _target(d)
        if not t:
            r = orig2(src, dst, *a, **k)
            _plain_copy(src, d)
            return r
        v = _carry(_before(t, _derived_source(src)), src, moved)
        r = orig2(src, dst, *a, **k)
        _after(t, v)
        return r

    def copy2(src, dst, *a, **k):
        return _copy2(False, src, dst, *a, **k)

    def copy2_moving(src, dst, *a, **k):      # move 가 복사 뒤 원본을 지울 때 쓴다 - 옮긴 사람 파일은 계속 사람 파일
        return _copy2(True, src, dst, *a, **k)

    m.copy2 = copy2
    if m.move.__defaults__ and m.move.__defaults__[-1] is orig2:
        m.move.__defaults__ = m.move.__defaults__[:-1] + (copy2_moving,)


def _patch_openpyxl(m):
    orig = m.ExcelReader.read

    def read(self):
        r = orig(self)
        _tag(self.wb, getattr(self.archive, "filename", None))
        return r

    m.ExcelReader.read = read


def _patch_opc(m):
    orig = m.OpcPackage.__dict__["open"].__func__

    def open(cls, pkg_file):
        pkg = orig(cls, pkg_file)
        _tag(pkg, pkg_file)
        return pkg

    m.OpcPackage.open = classmethod(open)


_COM_SAVE = ("Save", "SaveAs", "SaveAs2", "SaveCopyAs", "Close", "Quit")


def _com_plan(obj, attr, a, k):
    """COM 저장 호출이 쓸 문서 경로와 표식. 문서 저장이 아니면 None (메일 항목·PDF 저장 등은 그대로 지나간다)."""
    full = obj.FullName
    if not isinstance(full, str):
        return None
    name = None
    if attr == "Save":
        name = full
    elif attr == "Close":
        sc = a[0] if a else k.get("SaveChanges")
        if sc is not True and sc not in (1, -1):
            return None
        name = (a[1] if len(a) > 1 else k.get("Filename") or k.get("FileName")) or full
    else:
        name = a[0] if a else k.get("Filename") or k.get("FileName")
    if not isinstance(name, str) or not os.path.isabs(name):
        return None
    t = _target(os.path.abspath(name))
    if not t:
        return None
    g = _guard()
    if os.path.isabs(full) and g.same_path(t, full):
        return t, (g.source_of(t) if os.path.isfile(t) else None)   # 연 문서를 그 자리에 저장: 부분 편집
    return t, (_derived_source(full) if os.path.isabs(full) else None)


def _com_wrap(obj, attr, fn):
    def call(*a, **k):
        if attr == "Quit":
            r = fn(*a, **k)
            if _com_wait:
                _flush_com()
            return r
        full = None
        try:
            job = _com_plan(obj, attr, a, k)
            if attr == "Close":
                full = obj.FullName
        except Exception:
            job = None
        if not job:
            r = fn(*a, **k)
            if _com_wait and isinstance(full, str):
                w = _com_wait.get(os.path.normcase(os.path.abspath(full)))
                if w:
                    _after(*w)
            return r
        t, source = job
        v = _before(t, source)               # 사람 파일이면 여기서 멈춘다 - 엑셀이 덮어쓰기 전
        r = fn(*a, **k)
        if v is not None:
            _com_wait[os.path.normcase(os.path.abspath(t))] = (t, v)
            _after(t, v)                     # 열려 있어 못 읽으면 닫힐 때·끝날 때 다시 본다
        return r
    return call


def _patch_win32com(m):
    C = m.CDispatch
    orig = C.__getattr__

    def __getattr__(self, attr):
        r = orig(self, attr)
        if attr in _COM_SAVE and callable(r):
            try:
                return _com_wrap(self, attr, r)
            except Exception:
                return r
        return r

    C.__getattr__ = __getattr__


_PATCH = {
    "zipfile": _patch_zipfile,
    "shutil": _patch_shutil,
    "win32com.client.dynamic": _patch_win32com,
    "openpyxl.reader.excel": _patch_openpyxl,
    "docx.opc.package": _patch_opc,
    "pptx.opc.package": _patch_opc,
}


def _safely(fn, module):
    try:
        fn(module)
    except Exception:
        pass


class _Finder:
    """대상 모듈이 불려올 때만 실행 뒤에 끼어든다."""

    @classmethod
    def find_spec(cls, name, path=None, target=None):
        fn = _PATCH.get(name)
        if fn is None:
            return None
        spec = sys.modules["_frozen_importlib_external"].PathFinder.find_spec(name, path, target)
        if spec is None or spec.loader is None:
            return None
        orig = spec.loader.exec_module

        def exec_module(module):
            orig(module)
            _safely(fn, module)

        spec.loader.exec_module = exec_module
        return spec


def _install():
    _orig_replace = os.replace

    def replace(src, dst, *, src_dir_fd=None, dst_dir_fd=None):
        t = _target(dst)
        if not t:
            return _orig_replace(src, dst, src_dir_fd=src_dir_fd, dst_dir_fd=dst_dir_fd)
        v = _carry(_before(t, _derived_source(src)), src, True)
        r = _orig_replace(src, dst, src_dir_fd=src_dir_fd, dst_dir_fd=dst_dir_fd)
        _after(t, v)
        return r

    os.replace = replace

    _orig_rename = os.rename

    def rename(src, dst, *, src_dir_fd=None, dst_dir_fd=None):
        t = _target(dst)
        if not t or (os.name == "nt" and os.path.exists(t)):   # Windows 의 rename 은 기존 파일을 덮어쓰지 못한다
            return _orig_rename(src, dst, src_dir_fd=src_dir_fd, dst_dir_fd=dst_dir_fd)
        v = _carry(_before(t, _derived_source(src)), src, True)
        r = _orig_rename(src, dst, src_dir_fd=src_dir_fd, dst_dir_fd=dst_dir_fd)
        _after(t, v)
        return r

    os.rename = rename
    sys.meta_path.insert(0, _Finder)
    for name, fn in _PATCH.items():
        if name in sys.modules:
            _safely(fn, sys.modules[name])


try:
    _install()
except Exception:
    pass
