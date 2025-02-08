import InstrumentManager

class EtlFileManager:
    def __init__(self, instrument: InstrumentManager, folder_name: str):
        self.instrument = instrument
        self.folder_name = folder_name

    
    def 