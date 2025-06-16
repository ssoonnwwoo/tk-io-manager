class ShotModel:
    def __init__(self):
        self.sequence_name = ""
        self.shot_name = ""
        self.type = ""
        self.version = ""

    def get_sequence_name(self):
        return self.sequence_name

    def get_shot_name(self):
        return self.shot_name

    def get_type(self):
        return self.type

    def get_version(self):
        return self.version
    
    def set_sequence_name(self, value):
        self.sequence_name = value

    def set_shot_name(self, value):
        self.shot_name = value

    def set_type(self, value):
        self.type = value

    def set_version(self, value):
        self.version = value