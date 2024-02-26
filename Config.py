import yaml

class Config:
    def __init__(self, filename=None):
        inputFilename = filename if filename is not None else 'config.yaml'
        config = yaml.safe_load(open(inputFilename))
        self.data = config
        self.filename = inputFilename # for updating later

    # general set method
    def set(self, key, value):
        self.data[key] = value

    # general write method
    def save(self, filename=None):
        outputFilename = filename if filename is not None else self.filename
        yaml.safe_dump(self.data, open(outputFilename, 'w'), default_flow_style=False)

    # specific set for a common use case
    def setNextInvoiceNumber(self, value):
        self.set('nextInvoiceNumber', value)
        self.save()

if __name__ == '__main__':
    import sys

    inputFilename = sys.argv[1] if len(sys.argv) > 1 else None
    
    config = Config(filename=inputFilename)
    print(config.data)

    print(config.data['address']['line1'])
