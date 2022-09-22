import openpyxl as pyxl
import os
import pandas as pd
from datetime import datetime
from tqdm import tqdm


def basepath(dir: str = '') -> str:
    if os.name == "nt":
        BASEPATH = f"{os.getenv('USERPROFILE')}\\{dir}"
    else:
        BASEPATH = f"{os.getenv('HOME')}/{dir}"
    return BASEPATH


class ExcelMapper:
    __unit_no: int = 0;

    __target_headers: dict = {
        '0': 'voucher',
        '1' : 'journal_number',
        '2' : 'date',
        '4' : 'account',
        '8' : 'amount',
        '13' : 'remarks',
    }

    __mapping_headers: dict = {
        'date': 'Date',
        'unit_no': 'Unit No.',
        'd_division': 'DebitDivision',
        'd_project': 'DebitProject',
        'd_account': 'DebitAccount',
        'd_sub_account': 'DebitSub Account',
        'd_amount': 'DebitAmount',
        'd_remarks': 'DebitRemarks',
        'c_division': 'CreditDivision',
        'c_project': 'CreditProject',
        'c_account': 'CreditAccount',
        'c_sub_account': 'CreditSub Account',
        'c_amount': 'CreditAmount',
        'c_remarks': 'CreditRemarks',
        'voucher': 'Voucher Label',
        'voucher_type': 'Voucher Type',
    }

    def __init__(self, target: str, mapper: str) -> None:
        self.mapper = mapper
        self.target = target

        delimiter: str = '\\' if os.name == "nt" else '/'
        filename: str = target.split(delimiter).pop()

        extension: str = ''

        for suffix in ['.xlsx', '.xls']:
            if filename.endswith(suffix):
                filename = filename.replace(suffix, '')
                extension = suffix
                break

        filename = f'{filename}_Mapper_Result_{int(datetime.now().strftime("%Y%m%d%H%M%S"))}{extension}'

        self.resultpath = basepath(f'Downloads{delimiter}{filename}')
    
    @property
    def mapper(self) -> pd.DataFrame:
        return self.__mapper
    
    @mapper.setter
    def mapper(self, newmapper: str) -> None:
        abspath, _ = self.__extract_abspath(newmapper)
        df = pd.read_excel(
            abspath,
            header=None,
            skiprows=[0],
            dtype={0: str, 3: str, 5: str}
        )
        self.__mapper = df[[0, 3, 5]].fillna('')
    
    @property
    def resultpath(self) -> str:
        return self.__resultpath
    
    @resultpath.setter
    def resultpath(self, newresultpath: str) -> None:
        self.__resultpath = os.path.abspath(newresultpath)

    @property
    def target(self) -> pd.DataFrame:
        return self.__target
    
    @target.setter
    def target(self, newtarget: str) -> None:
        abspath, _ = self.__extract_abspath(newtarget)
        wb = pyxl.load_workbook(abspath, read_only=True)
        ws = wb.worksheets[0]

        rows = ws.rows
        columns = list(map(lambda x: x, self.__target_headers.values()))
        next(rows)
        
        data = []
        for row in rows:
            record = {}
            cells = [cell.value for i, cell in enumerate(row) if str(i) in self.__target_headers]

            for key, cell in zip(columns, cells):
                if type(cell) is datetime:
                    value = cell.strftime('%Y%m%d')
                elif cell is None:
                    value = ''
                else:
                    value = cell
                record[key] = value
            data.append(record)
    
        self.__target = pd.DataFrame(data, columns=columns)

    def __extract_abspath(self, path: str) -> tuple[str, str]:
        """ 
        Extract path to absolute path by os system
        
        :param path: path of target
        :return tuple[absolute path: str, sheet name: str]
        """
        abspath = os.path.abspath(path)
        file = pd.ExcelFile(abspath)
        sheet_name = file.sheet_names[0]
        return abspath, sheet_name
    
    def __find_account(self, taccount: pd.Series) -> tuple[any, any]:
        """ 
        Find bridgenote account given target account
        
        :param taccount: pd.Series target account
        :return tuple[account: any, sub_account: any]
        """
        df = self.__mapper.loc[self.__mapper[0].isin(taccount)]
    
        if not df.empty:
            account, sub_account = df[[3, 5]].values[0]
            return account, sub_account
        
        return f'N/A: {taccount.str}', None

    def __generate_detail_columns(self, side: str):
        columns: list = ['division', 'project', 'account', 'sub_account', 'amount', 'remarks']
        
        return list(map(lambda x: f'{side}_{x}', columns))

    def __create_credit(self, record: pd.DataFrame) -> pd.DataFrame:
        credit: pd.DataFrame = record.sort_values(by=['amount', 'remarks'], ascending=[False, True])
        credit: pd.DataFrame = credit.reset_index(drop=True)
        credit: pd.DataFrame = credit.rename(columns={'amount': 'c_amount', 'remarks': 'c_remarks'})
        
        credit['c_amount'] = (-credit['c_amount'])
        c_acount: pd.Series = credit.pop('account')
        c_bn_account, c_bn_sub_account = self.__find_account(c_acount)
        
        credit.insert(len(credit.columns.values), 'c_division', '')
        credit.insert(len(credit.columns.values), 'c_project', '')
        credit.insert(len(credit.columns.values), 'c_account', c_bn_account)
        credit.insert(len(credit.columns.values), 'c_sub_account', c_bn_sub_account if c_bn_sub_account is not None else 'N/A')

        return credit.reindex(columns=self.__generate_detail_columns('c'))
    
    def __create_debit(self, record: pd.DataFrame) -> pd.DataFrame:
        debit: pd.DataFrame = record.sort_values(by=['amount', 'remarks'], ascending=[True, True])
        debit: pd.DataFrame = debit.reset_index(drop=True)
        debit: pd.DataFrame = debit.rename(columns={'amount': 'd_amount', 'remarks': 'd_remarks'})

        d_acount: pd.Series = debit.pop('account')
        d_bn_account, d_bn_sub_account = self.__find_account(d_acount)
        
        debit.insert(len(debit.columns.values), 'd_division', '')
        debit.insert(len(debit.columns.values), 'd_project', '')
        debit.insert(len(debit.columns.values), 'd_account', d_bn_account)
        debit.insert(len(debit.columns.values), 'd_sub_account', d_bn_sub_account if d_bn_sub_account is not None else 'N/A')

        return debit.reindex(columns=self.__generate_detail_columns('d'))

    def __create_journal_detail(self, record: pd.DataFrame) -> pd.DataFrame:
        self.__unit_no += 1
        columns: list = ['account', 'amount', 'remarks']
        columns_reindex: list = list(map(lambda x: x, self.__mapping_headers.keys()))

        foreign: pd.DataFrame = record[['voucher', 'journal_number', 'date']].drop_duplicates().reset_index(drop=True)
        foreign.pop('journal_number')
        foreign.insert(len(foreign.columns.values), 'unit_no', int(self.__unit_no))
        foreign.insert(len(foreign.columns.values), 'voucher_type', '')

        debit: pd.DataFrame = self.__create_debit(record[record['amount'] > 0][columns])
        credit: pd.DataFrame = self.__create_credit(record[record['amount'] < 0][columns])

        content: pd.DataFrame = pd.concat([debit, credit], axis=1).fillna('')
        journal: pd.DataFrame = pd.concat([foreign, content], axis=1).fillna(method='ffill')
        journal: pd.DataFrame = journal.reindex(columns=columns_reindex)
        journal: pd.DataFrame = journal.rename(columns=self.__mapping_headers)
    
        return journal
        
    
    def __create_journals(self, record: pd.DataFrame) -> pd.DataFrame:
        """ 
        Filtering account start with 6
        
        :param record: pd.DataFrame
        :return any
        """
        if record['account'].str.startswith('6').any():
            return self.__create_journal_detail(record)
    
    def execute(self) -> None:
        """ 
        Execute mapper the excel file
        
        :return None
        """
        tqdm.pandas(desc="Mapping account")
        df: pd.DataFrame = self.__target.groupby(['date', 'journal_number', 'voucher']).progress_apply(self.__create_journals).reset_index(drop=True)
        print(f'Creating excel mapping')
        df.to_excel(self.__resultpath, index=False)
        print(f'Success create excel mappin in ({self.__resultpath})')


if __name__ == '__main__':
    is_exists_target: bool = False
    is_exists_mapper: bool = False

    default_path_target: str = basepath('target.xlsx')
    default_path_mapper: str = basepath('mapper.xlsx')

    while not is_exists_target and not is_exists_mapper:
        target: str = input(f'Input realpath of target mapper ({default_path_target}): ')
        target: str = target if target else default_path_target
        print(f'Target mapper: {target}')
        mapper: str = input(f'Input realpath of file mapper ({default_path_mapper}): ')
        mapper: str = mapper if mapper else default_path_mapper
        print(f'File mapper: {mapper}')
        
        if not os.path.exists(target):
            print(f'File "{target}" not exists!!!')
        else:
            is_exists_target = True
        
        if not os.path.exists(mapper):
            print(f'File "{mapper}" not exists!!!')
        else:
            is_exists_mapper = True
    
    df = ExcelMapper(target=target, mapper=mapper)
    df.execute()
