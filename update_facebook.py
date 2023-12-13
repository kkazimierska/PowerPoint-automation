import settings

print(settings.INPUT_EXCEL)
import pandas as pd
def get_data(file: str)->pd.DataFrame:
    competencies = pd.read_excel(file)
    return competencies

