import os
import pyseq
import shutil

def rename_sequence(seq_path, new_path):
    os.makedirs(new_path, exist_ok=True)

    # new_path ex: /home/rapa/show/my_project/seq/S038/S038_0020/plate/org/v001
    parts = new_path.strip("/").split("/")
    seq_shot = parts[-4]  # "S038_0020"
    typ = parts[-2]       # "org"
    ver = parts[-1]       # "v001"

    for root, dirs, seqs in pyseq.walk(seq_path):
        for seq in seqs:
            ext = os.path.splitext(seq[0].path)[1]  # ex: ".exr"
            for frame in seq:
                if ext == ".exr":
                    frame_num = frame.frame
                    new_name = f"{seq_shot}_{typ}_{ver}.{frame_num}{ext}"
                elif ext == ".mov":
                    new_name = f"{seq_shot}_{typ}_{ver}{ext}"

                old_path = frame.path
                new_file_path = os.path.join(new_path, new_name)
                shutil.copy2(old_path, new_file_path)
                print(f"[OK] RENAMED: {old_path} >>>>> {new_file_path}")
            print(f"{seq} rename & copy complete")