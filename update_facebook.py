import settings
import pandas as pd

print(settings.INPUT_EXCEL)

def get_data(file: str)->pd.DataFrame:
    """ Read the data.

    :param file: team competency file
    :return: data frame with columns as 
    :rtype: pd.DataFrame
    """
    competencies = pd.read_excel(file)
    return competencies

