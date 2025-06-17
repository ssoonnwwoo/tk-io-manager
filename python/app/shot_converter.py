import pyseq
import os
import sgtk
import shutil
import subprocess
import tempfile



logger = sgtk.platform.get_logger(__name__)

class ShotConverter:
    def __init__(self):
        self.scandata_path = ""
        self.org_path = ""
        self.sg_path = ""
        self.jpg_path = ""
    
    def set_scandata_path(self, scandata_path):
        self.scandata_path = scandata_path
    
    def set_org_path(self, org_path):
        self.org_path = org_path
        os.makedirs(self.org_path, exist_ok=True)
        self.jpg_path = org_path+"_jpg"
        os.makedirs(self.jpg_path, exist_ok=True)

    def set_sg_path(self, sg_path):
        self.sg_path = sg_path
        os.makedirs(self.sg_path, exist_ok=True)

    def set_paths(self, scandata_path, org_path, sg_path):
        self.set_scandata_path(scandata_path)
        self.set_org_path(org_path)
        self.set_sg_path(sg_path)

    def get_ext(self):
        for root, dirs, seqs in pyseq.walk(self.scandata_path):
            ext = seqs[0].tail()
            return ext
    
    def validate_paths(self, paths):
        validate_result = True
        for path in paths:
            if not os.path.exists(path):
                logger.error(f"Not valid shot path: {path}")
                validate_result = False
                return validate_result
        return validate_result


    def rename_scandata(self):
        paths = [self.scandata_path, self.org_path]
        validate_paths_result = self.validate_paths(paths)
        if not validate_paths_result:
            return False

        file_name = self.make_file_name(self.org_path, "org")
        for root, dirs, seqs in pyseq.walk(self.scandata_path):
            for seq in seqs:
                converted_count = 0
                ext = seq.tail()
                scandata_name = seq.head().strip(".")
                for frame in seq:
                    if ext == ".exr":
                        frame_num = frame.frame
                        new_name = f"{file_name}.{frame_num}{ext}"
                    old_path = frame.path
                    new_file_path = os.path.join(self.org_path, new_name)
                    shutil.copy2(old_path, new_file_path)
                    converted_count += 1
                print(f"{scandata_name} rename to {file_name}\nTotal {converted_count} frames are rename & moved")
        return True


    def make_file_name(self, new_path, typ):
        parts = new_path.strip("/").split("/")
        seq_shot = parts[-4]  # "S038_0020"
        # typ = parts[-2]       # "org"
        ver = parts[-1]       # "v001"
        file_name = f"{seq_shot}_{typ}_{ver}"
        return file_name
    
    def exrs_to_jpgs(self):
        for root, dirs, seqs in pyseq.walk(self.org_path):
            for seq in seqs:
                ext = seq.tail()
                length = seq.length()
                # if ext != ".exr":
                #     continue
                converted_count = 0
                for frame in seq:
                    exr_path = frame.path
                    base_name = os.path.splitext(os.path.basename(exr_path))[0]
                    # S001_0010_org_v003.894887
                    jpg_file_name = base_name + ".jpg"
                    jpg_file_path = os.path.join(self.jpg_path, jpg_file_name)

                    success = self.exr_to_jpg(exr_path, jpg_file_path)
                    if success: 
                        converted_count += 1
                        print(f"{converted_count}/{length} EXR frames successfully converted to JPG")
                    else:
                        return False
                print(f"{converted_count} frames converting complete ({int(converted_count/length*100)}%)")
        return self.jpg_path
    def exr_to_jpg(self, input_exr_path, output_jpg_path):
        """
        ffmpeg을 사용해 EXR 파일을 JPG로 변환
        Args:
            input_exr_path: 원본 EXR 파일 경로
            output_jpg_path: 출력 JPG 파일 경로
        Returns:
            bool: 변환 성공 여부
        """

        if not os.path.isfile(input_exr_path):
            logger.error(f"EXR does not exist: {input_exr_path}")
            return False
        
        cmd = [
        "oiiotool",
        input_exr_path,
        "--colorconvert", "ACES", "sRGB",
        "-o", output_jpg_path
        ]
        subprocess.run(cmd, check=True)
        return True

    def jpgs_to_video(self, vformat='mp4'):
        src_dir = self.jpg_path
        dest_dir = self.sg_path

        for root, dirs, seqs in pyseq.walk(src_dir):
            for seq in seqs:
                # ext = seq.tail()
                # jpg_name = seq.head().strip(".")
                # if ext != ".jpg":
                #     continue
                
                start_frame = seq.start()
                padding_str = seq.format('%p')
                input_pattern = os.path.join(src_dir, f"{seq.head()}{padding_str}{seq.tail()}")
                new_filename = self.make_file_name(dest_dir, "plate")
                output_path = os.path.join(dest_dir, f"{new_filename}.{vformat}")

                # print(f"[INFO] Converting sequence: {input_pattern}")
                # print(f"[INFO] Output video: {output_path}")
                
                if vformat == "mp4":
                    codec = "libx264"
                    pix_fmt = "yuv420p"
                elif vformat == "webm":
                    codec = "libvpx"
                    pix_fmt = "yuv420p"
                else:
                    logger.error(f"Unsupported format: {vformat}")
                    return None

                cmd = [
                    "ffmpeg",
                    "-loglevel", "error",           
                    "-y",
                    "-start_number", str(start_frame),
                    "-framerate", "23.976",
                    "-i", input_pattern,
                    "-c:v", codec,
                    "-crf", "23",
                    "-pix_fmt", pix_fmt,
                    output_path,
                ]


                subprocess.run(cmd, check=True)
            print(f"{vformat.upper()} successfully created: {new_filename}")
        return output_path

    def jpgs_to_montage(self):
        src_path = self.jpg_path
        dest_dir = self.sg_path
        frame_increment = 5
        frame_width = 240
        fps = 23.976
        for _, _, seqs in pyseq.walk(src_path):
            for seq in seqs:
                _, ext = os.path.splitext(seq.name)
                # if ext != ".exr":
                #     continue

                head = seq.head()
                tail = seq.tail()
                start = seq.start()
                padding_str = seq.format('%p')
                input_pattern = os.path.join(src_path, f"{head}{padding_str}{tail}")

                with tempfile.TemporaryDirectory(prefix="filmstrip_") as tmp_dir:
                    tmp_pattern = os.path.join(tmp_dir, "thumb_%02d.jpeg")
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
                    #ImageMagick
                    subprocess.run(extract_cmd, check=True)
                    montage_name = self.make_file_name(self.sg_path, "montage") + ".jpg"
                    output_dir = os.path.join(dest_dir, montage_name)
                    concat_cmd = [
                        "convert",
                        "+append",
                        os.path.join(tmp_dir, "thumb_*.jpeg"),
                        output_dir
                    ]
                    subprocess.run(" ".join(concat_cmd), check=True, shell=True)
                    print(f"MONTAGE successfully created: {montage_name}")
                    return output_dir
        return False
    
    def jpgs_to_thumbnail(self):
        src_path = self.jpg_path
        dest_dir = self.sg_path

        thumbnail_name = self.make_file_name(self.sg_path,"thumbnail") + ".jpg"

        for _, _, seqs in pyseq.walk(src_path):
            for seq in seqs:
                head = seq.head()
                tail = seq.tail()
                start = seq.start()
                padding_str = seq.format('%p')
                # ex : "%07d" % 1 >>> "0000001"
                first_frame = padding_str % start
                input_path = os.path.join(src_path, f"{head}{first_frame}{tail}")
                output_path = os.path.join(dest_dir, thumbnail_name)
                print(f"THUMBNAIL successfully created: {thumbnail_name}")
                shutil.copy2(input_path, output_path)
                return output_path
        return False
