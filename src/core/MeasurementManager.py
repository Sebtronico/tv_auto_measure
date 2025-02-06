import InstrumentManager

class EtlMeasurementManager:
    def __init__(self, instrument: InstrumentManager = None):
        self.instrument = instrument