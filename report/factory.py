import credential
import pymssql
from report.delivery import ford


REPORT_TYPES_DICT = {
    'ford_inc_equip': ford.EquipmentAndIncentives,
}

REPORT_VIEWS_DICT = {
    'ford_inc_equip': 'select * from vw_ford_incentive_equipment where sample_date between {0} and {1}'
                      ' order by make, model, version, production_year, model_year, code',
}


def create_report(report_name: str, min_date: int, max_date: int):
    def retrieve_data():
        cred = credential.get_credential('ukvsqlbdrep01')
        with pymssql.connect(cred['address'], cred['username'], cred['password'], 'rt') as conn:
            with conn.cursor() as cursor:
                cursor.execute(REPORT_VIEWS_DICT[report_name].format(str(min_date), str(max_date)))
                return cursor.fetchall()

    return REPORT_TYPES_DICT[report_name](retrieve_data())
