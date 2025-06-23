import subprocess
import json
def get_timecode(exr_path):
    result = subprocess.run(
        ["exiftool", "-json", "-TimeCode", exr_path],
        capture_output=True,
        text=True
        )
    data = json.loads(result.stdout)[0]
    raw = data.get("TimeCode", "")
    if not raw:
        result = ""
        return result
    raw = int(raw.split(" ")[0].strip())
    
    # frame
    fu =  raw        & 0x0F          # 프레임 단위 자리 (bits 0–3)
    ft = (raw >> 4)  & 0x03          # 프레임 십의 자리 (bits 4–5)
    frames = ft * 10 + fu

    # seconds
    su = (raw >> 8)  & 0x0F          # 초 단위 자리 (bits 8–11)
    st = (raw >> 12) & 0x07          # 초 십의 자리 (bits 12–14)
    seconds = st * 10 + su

    # minutes
    mu = (raw >> 16) & 0x0F          # 분 단위 자리 (bits 16–19)
    mt = (raw >> 20) & 0x07          # 분 십의 자리 (bits 20–22)
    minutes = mt * 10 + mu

    # hours
    hu = (raw >> 24) & 0x0F          # 시 단위 자리 (bits 24–27)
    ht = (raw >> 28) & 0x03          # 시 십의 자리 (bits 28–29)
    hours = ht * 10 + hu

    timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frames:02d}"
    print(timecode)
    return timecode
