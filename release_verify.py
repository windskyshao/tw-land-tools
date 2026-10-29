"""更新檔簽章驗證(Ed25519，RFC 8032)：發布時用只有開發者電腦有的私鑰簽名(sign_release.py)，
主程式／地籍圖處理工具／電子謄本結構化下載更新後，用這裡內建的公鑰驗證，對不上就不安裝。
就算 GitHub 帳號被盜，對方也做不出能通過驗證的假版本。純 Python、不需要額外套件。
簽的內容 = 「檔名\\n」+ 檔案 sha256；簽名檔 = 檔名.sig(base64)，跟檔案一起放在 GitHub release。"""
import base64
import hashlib

# 發布用公鑰(私鑰在開發者電腦 %USERPROFILE%\.tw_release\release_signing_key.pem，不進版控、不打包)
PUBLIC_KEY_B64 = "LIdLQoF7KfHYF2+Fmr6I1axXv/AbHEE/1YjYxzKiok8="

_p = 2 ** 255 - 19
_q = 2 ** 252 + 27742317777372353535851937790883648493
_d = -121665 * pow(121666, _p - 2, _p) % _p
_SQRT_M1 = pow(2, (_p - 1) // 4, _p)


def _add(P, Q):
    A = (P[1] - P[0]) * (Q[1] - Q[0]) % _p
    B = (P[1] + P[0]) * (Q[1] + Q[0]) % _p
    C = 2 * P[3] * Q[3] * _d % _p
    D = 2 * P[2] * Q[2] % _p
    E, F, G, H = B - A, D - C, D + C, B + A
    return (E * F, G * H, F * G, E * H)


def _mul(s, P):
    Q = (0, 1, 1, 0)
    while s > 0:
        if s & 1:
            Q = _add(Q, P)
        P = _add(P, P)
        s >>= 1
    return Q


def _eq(P, Q):
    return (P[0] * Q[2] - Q[0] * P[2]) % _p == 0 and (P[1] * Q[2] - Q[1] * P[2]) % _p == 0


def _recover_x(y, sign):
    if y >= _p:
        return None
    x2 = (y * y - 1) * pow(_d * y * y + 1, _p - 2, _p)
    if x2 == 0:
        return None if sign else 0
    x = pow(x2, (_p + 3) // 8, _p)
    if (x * x - x2) % _p != 0:
        x = x * _SQRT_M1 % _p
    if (x * x - x2) % _p != 0:
        return None
    if (x & 1) != sign:
        x = _p - x
    return x


_gy = 4 * pow(5, _p - 2, _p) % _p
_gx = _recover_x(_gy, 0)
_G = (_gx, _gy, 1, _gx * _gy % _p)


def _decompress(s):
    if len(s) != 32:
        return None
    y = int.from_bytes(s, "little")
    sign = y >> 255
    y &= (1 << 255) - 1
    x = _recover_x(y, sign)
    return None if x is None else (x, y, 1, x * y % _p)


def ed25519_verify(public, msg, sig):
    if len(public) != 32 or len(sig) != 64:
        return False
    A = _decompress(public)
    R = _decompress(sig[:32])
    if not A or not R:
        return False
    s = int.from_bytes(sig[32:], "little")
    if s >= _q:
        return False
    h = int.from_bytes(hashlib.sha512(sig[:32] + public + msg).digest(), "little") % _q
    return _eq(_mul(s, _G), _add(R, _mul(h, A)))


def file_message(path, name):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return name.encode("utf-8") + b"\n" + h.digest()


def verify_file(path, name, sig_b64):
    """path＝下載的檔案、name＝GitHub 上的檔名、sig_b64＝檔名.sig 的內容。通過回 True"""
    try:
        return ed25519_verify(base64.b64decode(PUBLIC_KEY_B64), file_message(path, name),
                              base64.b64decode(sig_b64.strip()))
    except Exception:
        return False
