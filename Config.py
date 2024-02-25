import yaml

class Config:
    def __init__(self, filename=None):
        inputFilename = filename if filename is not None else 'config.yaml'
        config = yaml.safe_load(open(inputFilename))
        self.data = config

if __name__ == '__main__':
    import sys

    inputFilename = sys.argv[1] if len(sys.argv) > 1 else None
    
    config = Config(filename=inputFilename)
    print(config.data)

    print(config.data['address']['line1'])
