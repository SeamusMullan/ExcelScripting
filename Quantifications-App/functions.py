import pandas

logging = True

class logColors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    


def log(msg, status='log'):
    """Logs messages to console with a category

    Args:
        msg (str): message to be logged
        status (str, optional): category of the log, either `log`, `message`, `warning` or `error`. Defaults to `log`.
    """
    if (logging):
        match status:
            case 'log':
                print(logColors.OKBLUE + "LOG: " + msg + logColors.ENDC)
                
            case 'message':
                print(logColors.OKGREEN + "MESSAGE: " + msg + logColors.ENDC)
            
            case 'warning':
                print(logColors.WARNING + "WARNING: " + msg + logColors.ENDC)
                
            case 'error':
                print(logColors.FAIL + "ERROR: " + msg + logColors.ENDC)
                

def convert_csv_file(fileName, outputName="output_file"):
    """Converts a csv file to an xlsx file

    Args:
        fileName (str): Name of csv file
        outputName (str, optional): Name of output file. Defaults to "output_file".
    """
    read_file = pandas.read_csv(fileName)
    read_file.to_excel(f'{outputName}.xlsx', index=None, header=True)
    

def removeControlCharacters(input):
    """Filters control characters from a string using regex
    

    Args:
        input (str): The Input string

    Returns:
        str: The filtered string
    """
    controlCharsRegex = "/[\u0000-\u001F\u007F-\u009F]/g"
    return input.replace(controlCharsRegex, '')


def merge_columns(columns):
    """
    Merge two columns represented as `pandas.Series` objects into one

    Args:
        `pandas.Series[]`: Array of columns

    Returns:
        `pandas.Series`: Merged column of data
    """

    mergedColumn = pandas.Series([])
    
    for col in columns:
        pandas.merge_ordered(mergedColumn, col)
    
    return mergedColumn


def updateNamesFromTable(col, table):
    pass