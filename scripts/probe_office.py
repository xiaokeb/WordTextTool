"""Read local Office automation identity and required capability flags."""
import json
import winreg


def registrations():
    found = {}
    for progid in ("Excel.Application", "ket.Application"):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + "\\CLSID") as key:
                clsid = winreg.QueryValue(key, None)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "CLSID\\" + clsid + "\\LocalServer32") as key:
                server = winreg.QueryValue(key, None)
            found[progid] = {"clsid": clsid, "server": server}
        except OSError as exc:
            found[progid] = {"error": str(exc)}
    return found


if __name__ == "__main__":
    print(json.dumps(registrations(), ensure_ascii=False, indent=2))
