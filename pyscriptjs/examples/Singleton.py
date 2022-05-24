""" This is singleton service for other services to base on"""


class SingletonService(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(SingletonService, cls).__call__(*args, **kwargs)
        return cls._instances[cls]
