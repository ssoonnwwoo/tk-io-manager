import subprocess
import os
import pyseq
import tempfile

def exr_to_jpg(input_exr_path, output_jpg_path):
    """
    ffmpeg을 사용해 EXR 파일을 JPG로 변환
    Args:
        input_exr_path: 원본 EXR 파일 경로
        output_jpg_path: 출력 JPG 파일 경로
    Returns:
        bool: 변환 성공 여부
    """

    if not os.path.isfile(input_exr_path):
        print(f"[ERROR] EXR does not exist: {input_exr_path}")
        return False

    try:
        cmd = ["ffmpeg", "-loglevel", "error", "-y", "-i", input_exr_path, "-q:v", "2", output_jpg_path]
        subprocess.run(cmd, check=True)
        print(f"[COMPLETE] Input : {input_exr_path}")
        print(f"[COMPLETE] Output : {output_jpg_path}")
        return True

    # subprocess에서 run 실패
    except Exception as e:
        print(f"[EXCEPTION] Error occurred while attempting to convert thumbnail: {e}")
        return False


def mov_to_jpg(input_mov_path, output_jpg_path, all_frames=False):
    """
    ffmpeg을 사용해 MOV 파일에서 JPG 이미지 추출
    Args:
        input_mov_path: 원본 MOV 파일 경로
        output_jpg_path: 출력 JPG 경로 (첫 프레임일 경우는 파일, 시퀀스일 경우는 폴더/image_%04d.jpg 형식)
        all_frames (bool): True면 전체 프레임 시퀀스 저장, False면 첫 프레임만 저장
    Returns:
        bool: 추출 성공 여부
    """

    if not os.path.isfile(input_mov_path):
        print(f"[ERROR] MOV does not exist: {input_mov_path}")
        return False

    try:
        cmd = ["ffmpeg", "-loglevel", "error", "-y", "-i", input_mov_path]
        
        if all_frames:
            # 전체 프레임 추출
            os.makedirs(os.path.dirname(output_jpg_path), exist_ok=True)
            cmd += ["-q:v", "2", os.path.join(output_jpg_path, "image_%04d.jpg")]
        else:
            # 첫 프레임만 추출
            cmd += ["-frames:v", "1", "-q:v", "2", output_jpg_path]

        subprocess.run(cmd, check=True)
        print(f"[COMPLETE] Input : {input_mov_path}")
        print(f"[COMPLETE] Output : {output_jpg_path}")
        return True

    except Exception as e:
        print(f"[EXCEPTION] Error occurred while extracting JPG from MOV: {e}")
        return False

def mov_to_exrs(mov_path, output_dir):
    parts = output_dir.strip("/").split("/")
    seq_shot = parts[-4]  # "S038_0020"
    typ = parts[-2]       # "org"
    ver = parts[-1]       # "v001"
    # S040_0020_org_v001.######.exr
    
    output_pattern = os.path.join(
        output_dir, f"{seq_shot}_{typ}_{ver}.%7d.exr"
    )

    try:
        ffmpeg_path = "/opt/ffmpeg-openexr/bin/ffmpeg"
        cmd = [ffmpeg_path, "-loglevel", "error", "-y", "-i", mov_path, "-c:v", "exr", output_pattern]
        subprocess.run(cmd, check=True)
        print(f"[COMPLETE] MOV to EXR: {output_pattern}")
        return True
    except Exception as e:
        print(f"[EXCEPTION] Error converting MOV to EXR: {e}")
        return False


def exrs_to_jpgs(src_dir, dest_dir):
    if not os.path.isdir(src_dir):
        print(f"[ERROR] Source directory does not exist: {src_dir}")
        return False
    

    for root, dirs, seqs in pyseq.walk(src_dir):
        for seq in seqs:
            _, ext = os.path.splitext(seq.name)
            converted_count = 0
            if  ext != ".exr":
                continue

            print(f"\n[INFO] Processing sequence: {seq.head()}{seq.tail()} ({len(seq)} frames)")

            for frame in seq:
                src_path = frame.path
                base_name = os.path.splitext(os.path.basename(src_path))[0]
                jpg_name = base_name + ".jpg"
                jpg_path = os.path.join(dest_dir, jpg_name)

                success = exr_to_jpg(src_path, jpg_path)
                if success:
                    converted_count += 1
            print(f"\n[COMPLETE] {converted_count} EXR frames successfully converted to JPG.")
    return True

def exrs_to_video(src_dir, dest_dir, vformat='mp4'):
    parts = dest_dir.strip("/").split("/")
    seq_shot = parts[-4]  # "S038_0020"
    ver = parts[-1]       # "v001"

    if not os.path.isdir(src_dir):
        print(f"[ERROR] Source directory does not exist: {src_dir}")
        return None

    for root, dirs, seqs in pyseq.walk(src_dir):
        for seq in seqs:
            _, ext = os.path.splitext(seq.name)
            if ext != ".exr":
                continue
            
            start_frame = seq.start()
            padding_str = seq.format('%p')
            input_pattern = os.path.join(src_dir, f"{seq.head()}{padding_str}{seq.tail()}")
            output_path = os.path.join(dest_dir, f"{seq_shot}_plate_{ver}.{vformat}")

            # print(f"[INFO] Converting sequence: {input_pattern}")
            # print(f"[INFO] Output video: {output_path}")

            if vformat == "mp4":
                codec = "libx264"
                pix_fmt = "yuv420p"
            elif vformat == "webm":
                codec = "libvpx"
                pix_fmt = "yuv420p"
            else:
                print(f"[ERROR] Unsupported format: {vformat}")
                return None

            # ffmpeg -framerate 23.976 -i input.%07d.exr -c:v libx264 -crf 25 -pix_fmt yuv420p output.mp4
            cmd = [
                "ffmpeg",
                "-loglevel", "error",           # 오류만 출력
                "-y",
                "-start_number", str(start_frame),
                "-framerate", "23.976",
                "-i", input_pattern,
                "-c:v", codec,
                "-crf", "25",
                "-pix_fmt", pix_fmt,
                output_path,
            ]

            try:
                subprocess.run(cmd, check=True)
                print(f"[COMPLETE] {vformat.upper()} successfully created at {output_path}")
            except subprocess.CalledProcessError as e:
                print(f"[ERROR] FFmpeg failed: {e}")
                return None

    return output_path

def exrs_to_montage(src_path, dest_dir):
    parts = dest_dir.strip("/").split("/")
    seq_shot = parts[-4]  # "S038_0020"
    ver = parts[-1]       # "v001"
    frame_increment = 5
    frame_width = 240
    fps = 23.976

    for _, _, seqs in pyseq.walk(src_path):
        for seq in seqs:
            _, ext = os.path.splitext(seq.name)
            if ext != ".exr":
                continue

            head = seq.head()
            tail = seq.tail()
            start = seq.start()
            padding_str = seq.format('%p')
            input_pattern = os.path.join(src_path, f"{head}{padding_str}{tail}")

            with tempfile.TemporaryDirectory(prefix="filmstrip_") as tmp_dir:
                tmp_pattern = os.path.join(tmp_dir, "thumb_%02d.jpeg")

                try:
                    extract_cmd = [
                        "ffmpeg",
                        "-loglevel", "error",
                        "-start_number", str(start),
                        "-i", input_pattern,
                        "-vf", f"select='not(mod((n-{start})\\,{frame_increment}))',setpts='N/({fps}*TB)',scale={frame_width}:-1",
                        "-sws_flags", "lanczos",
                        "-qscale:v", "2",
                        "-pix_fmt", "yuvj420p",
                        "-f", "image2",
                        tmp_pattern
                    ]
                    subprocess.run(extract_cmd, check=True)

                    montage_name = f"{seq_shot}_montage_{ver}.jpg"
                    output_dir = os.path.join(dest_dir, montage_name)
                    concat_cmd = [
                        "convert",
                        "+append",
                        os.path.join(tmp_dir, "thumb_*.jpeg"),
                        output_dir
                    ]
                    subprocess.run(" ".join(concat_cmd), check=True, shell=True)

                    print(f"[COMPLETE] Montage created at {dest_dir}")
                    return output_dir

                except subprocess.CalledProcessError as e:
                    print(f"[ERROR] Failed to generate montage: {e}")
                    return None

    print("[ERROR] No EXR sequences found.")
    return None

def exrs_to_thumbnail(src_path, dest_dir):
    parts = dest_dir.strip("/").split("/")
    seq_shot = parts[-4]  # "S038_0020"
    ver = parts[-1]       # "v001"

    thumbnail_name = f"{seq_shot}_thumbnail_{ver}.jpg"
    output_path = os.path.join(dest_dir, thumbnail_name)

    for _, _, seqs in pyseq.walk(src_path):
        for seq in seqs:
            _, ext = os.path.splitext(seq.name)
            if ext != ".exr":
                continue

            head = seq.head()
            tail = seq.tail()
            start = seq.start()
            padding_str = seq.format('%p')
            # ex : "%07d" % 1 >>> "0000001"
            first_frame = padding_str % start
            input_path = os.path.join(src_path, f"{head}{first_frame}{tail}")
            if exr_to_jpg(input_path, output_path):
                print(f"[COMPLETE] Thumbnail created at {output_path}")
                return output_path
            else:
                print(f"[ERROR] Failed to create thumbnail for {input_path}")
                return None
            
    print("[ERROR] No EXR sequences found.")
    return None

