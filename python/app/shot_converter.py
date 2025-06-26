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
        self.fps = 0.0
        self.start_frame = 1001
    
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

    def set_fps(self, fps):
        self.fps = fps

    # def set_start_frame(self, start_frame):
    #     self.start_frame = start_frame
    

    def set_paths(self, scandata_path, org_path, sg_path, fps):
        self.set_scandata_path(scandata_path)
        self.set_org_path(org_path)
        self.set_sg_path(sg_path)
        self.set_fps(fps)
        # self.set_start_frame(start_frame)

    def get_ext(self):
        for root, dirs, seqs in pyseq.walk(self.scandata_path):
            ext = seqs[0].tail()
            return ext
        
    def make_sg_datas(self):
        paths = [self.scandata_path, self.org_path, self.jpg_path, self.sg_path]
        if not self.validate_paths(paths):
            return False
        
        _, _, seqs = next(pyseq.walk(self.scandata_path))
        seq = seqs[0]
        result = self._process_start(seq)
        return result

    def _process_start(self, seq):
        """"
        Rename & Copy ->
        EXRs to JPGs ->
        EXRs to MOV ->
        JPGs to Video ->
        make thumbnail & montage
        """
        scandata_name = seq.head().strip(".")
        scandata_length = seq.length()
        # Rename & Copy
        file_base = self.make_file_name(self.org_path, "org") # S001_0010_org_v001
        start_frame = int(self.start_frame)
        copied_count = 0
        org_files = []
        for frame in seq:
            new_name = f"{file_base}.{str(start_frame)}{seq.tail()}" # S001_0010_org_v001.1001.exr
            dst_path = os.path.join(self.org_path, new_name)
            shutil.copyfile(frame.path, dst_path)
            org_files.append(dst_path)
            start_frame += 1
            copied_count += 1
        print(f"{scandata_name} rename to {file_base}\nTotal {copied_count}/{scandata_length} frames are rename & copied")
        
        # EXRs to JPGs
        jpg_paths = []
        renamed_seq_length = len(org_files)
        converted_count = 0
        for exr_path in org_files:
            base = os.path.splitext(os.path.basename(exr_path))[0]
            jpg_name = base + ".jpg" # S001_0010_org_v001.1001.jpg
            jpg_path = os.path.join(self.jpg_path, jpg_name)
            if self.exr_to_jpg(exr_path, jpg_path):
                jpg_paths.append(jpg_path)
                converted_count += 1 
                print(f"{converted_count}/{renamed_seq_length} EXR frames successfully converted to JPG")
            else:
                logger.error(f"Error occurred while coverting exrs to jpgs")
                result_paths = {"mp4" : None, "webm" : None, "montage" : None, "thumbnail" : None}
                return result_paths
        
        mp4 = self._seq_to_video(jpg_paths, vformat="mp4")
        webm = self._seq_to_video(jpg_paths, vformat="webm")
        mov = self._seq_to_video(jpg_paths, vformat="mov")
        montage = self._jpgs_to_montage(jpg_paths)
        thumbnail = self._jpgs_to_thumbnail(jpg_paths)

        result_paths = {"mp4" : mp4, "webm" : webm, "montage" : montage, "thumbnail" : thumbnail}
        return result_paths

    def validate_paths(self, paths):
        validate_result = True
        for path in paths:
            if not os.path.exists(path):
                logger.error(f"Not valid shot path: {path}")
                validate_result = False
                return validate_result
        return validate_result

    def make_file_name(self, new_path, typ):
        """
        Make new file name from org path or jpg path
        
        Args:
            new_path (str): org path or jpg path(.../seq/S001/S001_0010/plate/org/v001)
            typ (str): org, plate, ...
        
        Returns:
            file_name (str): new file name (S001_0010_org_v001)
        """
        parts = new_path.strip("/").split("/")
        seq_shot = parts[-4]  # "S038_0020"
        ver = parts[-1]       # "v001"
        file_name = f"{seq_shot}_{typ}_{ver}"
        return file_name

    def exr_to_jpg(self, input_exr_path, output_jpg_path):
        if not os.path.isfile(input_exr_path):
            logger.error(f"EXR does not exist: {input_exr_path}")
            return False
        
        cmd = [
        "oiiotool",
        input_exr_path,
        "--colorconvert", "ACES", "sRGB",
        "-o", output_jpg_path
        ]
        result = subprocess.run(cmd)
        if result.returncode != 0:
            logger.error(f"JPG convert failed. {output_jpg_path}")
            return False
        return True

    def _seq_to_video(self, seq, vformat='mp4'):
        first = seq[0]
        dirname, fname = os.path.split(first) 
        base, _ = os.path.splitext(fname) # "S038_0020_org_v003.1001"
        prefix, num = base.rsplit('.', 1) # prefix="S038_0020_org_v003", num="1001"
        start_frame = int(num)
        input_pattern = os.path.join(dirname, f"{prefix}.%d.jpg")
        fps = round(float(self.fps), 4)
        fps = f"{fps:.3f}"
        output_name = self.make_file_name(self.sg_path, 'plate') + f".{vformat}"
        output_path = os.path.join(self.sg_path, output_name)
        
        cmd = [
            "ffmpeg",
            "-loglevel", "error",
            "-y",
            "-start_number", str(start_frame),
            "-framerate", fps,
            "-f", "image2",
            "-i", input_pattern,
        ]

        if vformat == "mp4":
            # H.264, 8bit
            cmd += [
                "-c:v", "libx264",
                "-crf", "23",
                "-pix_fmt", "yuv420p",
            ]
        elif vformat == "webm":
            # VP8/VP9, 8bit
            cmd += [
                "-c:v", "libvpx",
                "-crf", "23",
                "-pix_fmt", "yuv420p",
            ]

        elif vformat == "mov":
            cmd += [
                "-vf", "eq=gamma=2.2,format=yuv422p10le",
                "-c:v", "prores_ks",
                "-profile:v", "3",           
                "-pix_fmt", "yuv422p10le",
                "-color_primaries", "bt709",
                "-color_trc",      "bt709",
                "-colorspace",     "bt709",
            ]

        else:
            logger.error(f"Unsupported format: {vformat}")
            return False
        
        cmd.append(output_path)

        result = subprocess.run(cmd)
        if result.returncode != 0:
            logger.error(f"{vformat.upper()} created failed: {output_name}")
            return False
        
        print(f"{vformat.upper()} successfully created: {output_name}")
        return output_path
    
    def _jpgs_to_montage(self, jpg_paths):
        first = jpg_paths[0]
        jpg_dir = os.path.dirname(first)
        fname = os.path.basename(first) 
        base, ext = os.path.splitext(fname) # base: "S038_0020_org_v003.1001", ext: ".jpg"
        prefix, num = base.rsplit('.', 1) # prefix: "S038_0020_org_v003", num: "1001"
        start_frame = int(num)

        input_pattern = os.path.join(jpg_dir, f"{prefix}.%d{ext}")

        frame_increment = 5
        frame_width = 240
        fps = self.fps

        with tempfile.TemporaryDirectory(prefix="filmstrip_") as tmp_dir:
            tmp_pattern = os.path.join(tmp_dir, "thumb_%02d.jpeg")
            extract_cmd = [
                        "ffmpeg",
                        "-loglevel", "error",
                        "-start_number", str(start_frame),
                        "-i", input_pattern,
                        "-vf", f"select='not(mod((n-{start_frame})\\,{frame_increment}))',setpts='N/({fps}*TB)',scale={frame_width}:-1",
                        "-sws_flags", "lanczos",
                        "-qscale:v", "2",
                        "-pix_fmt", "yuvj420p",
                        "-f", "image2",
                        tmp_pattern
                    ]
            result = subprocess.run(extract_cmd)
            if result.returncode != 0:
                logger.error(f"Error occurred while extracting jpgs for flimstrip")
                print(f"Error occurred while extracting jpgs for flimstrip")
                return False
            
            # ImageMagick
            montage_name = self.make_file_name(self.sg_path, "montage") + ".jpg"
            output_path  = os.path.join(self.sg_path, montage_name)
            concat_cmd = [
                "convert",
                "+append",
                os.path.join(tmp_dir, "thumb_*.jpeg"),
                output_path
            ]
            subprocess.run(" ".join(concat_cmd), shell=True)
            if result.returncode != 0:
                logger.error(f"Error occurred while making flimstrip jpg")
                print(f"Error occurred while making flimstrip jpg")
                return False
            print(f"MONTAGE successfully created: {montage_name}")
            return output_path

    def _jpgs_to_thumbnail(self, jpg_paths):
        thumbnail_name = self.make_file_name(self.sg_path, 'thumbnail') + '.jpg'
        output_path    = os.path.join(self.sg_path, thumbnail_name)
        first_jpg = jpg_paths[0]
        shutil.copyfile(first_jpg, output_path)
        print(f"THUMBNAIL successfully created: {thumbnail_name}")
        return output_path
