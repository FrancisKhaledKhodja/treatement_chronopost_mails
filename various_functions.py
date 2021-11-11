from time import time
import re

def time_execution(treatment_title: str):
    """decorator measuring the execution time of a function

    Args:
        treatment_title (str): Name of the function, for instance.
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            print("{0}\n{1}\n{0}\n".format("*" * len(treatment_title), treatment_title))
            t0 = time()
            func(*args, **kwargs)
            print("Durée du traitement: {} secondes".format(time() - t0))
            print("{0}\nFIN DE {1}\n{0}\n".format("*" * len(treatment_title), treatment_title))
        return wrapper
    return decorator


def format_date(date_to_format):
    motif = re.compile("[0-9]{8}")
    if not(isinstance(date_to_format, float)) and motif.search(date_to_format):
        return "{}/{}/{}".format(date_to_format[6:], date_to_format[4:6], date_to_format[:4])
    else:
        return date_to_format