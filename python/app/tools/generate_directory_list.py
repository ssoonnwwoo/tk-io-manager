def generate_directory_list(meta_list):
    directories = []
    for meta in meta_list:
        directory = meta.get("Directory")
        if directory:
            directories.append(directory)
    return directories
