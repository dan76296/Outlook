import logging

class Log():

    def __init__(self):
        _logger = logging.getLogger(__name__)
        _logger.setLevel(logging.DEBUG)
        # create file handler which logs even debug messages
        fh = logging.FileHandler('outlook.log')
        fh.setLevel(logging.DEBUG)
        # create console handler with a higher log level
        ch = logging.StreamHandler()
        ch.setLevel(logging.ERROR)
        # create formatter and add it to the handlers
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        # add the handlers to the logger
        _logger.addHandler(fh)
        _logger.addHandler(ch)

    def error(self, message):
        self.logging.error(message)
        pass

    def info(self, message):
        self.logging.info(message)
        pass
