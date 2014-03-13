# coding: utf-8
import xlrd
import logging


log = logging.getLogger(__name__)


def cell_value(cell):
    try:
        return unicode(cell.value)
    except:
        log.error(
            'cell_value - can not convert to unicode {!r}'
            .format(cell and cell.value)
        )
        raise


def cell_values(cell_seq):
    return [cell_value(cell) for cell in cell_seq]


def to_int(str):
    try:
        return int(float(str))
    except:
        log.error('to_int: can not convert: {!r}, substituting 0'.format(str))
        return 0


class Field(object):

    def __init__(self, name, description, length, type):
        self.name = name
        self.description = description
        self.length = length
        KNOWN_TYPES = (None, '', 'none', 'date', 'prop', 'num', 'char')
        assert type in KNOWN_TYPES, 'unknown type: <{}>'.format(type)
        self.type = type


# primary key
FIELD_CEG_ID = Field(u'ceg_id', u'Complex cég azonosító', 11, u'char')
FIELD_ALROVAT_ID = Field(u'alrovat_id', u'Alrovat', 8, u'num')


class Table(object):

    def __init__(self, name, description):
        self.name = name
        self.description = description
        self.fields = [FIELD_CEG_ID, FIELD_ALROVAT_ID]

    def add(self, field):
        self.fields.append(field)


def read_tables(r_export_txt_xls='R_export.txt.xls'):
    '''
    Create Complex table definitions directly from Complex provided xls file.
    '''
    wb = xlrd.open_workbook(r_export_txt_xls)
    assert wb.sheet_names() == [u'R_export']
    sheet = wb.sheets()[0]

    header = cell_values(sheet.row(0)[:6])
    assert header[0] == u'#Rovat'
    assert header[4:6] == [u'Hossz', u'IndexTipus']

    tables = []

    for irow in range(1, sheet.nrows):
        row = [unicode(v.value) for v in sheet.row(irow)[:6]]
        table_id, _, field_desc, field_name, field_length, field_type = row

        if field_length:
            field_length = to_int(field_length)
        if field_type in {u'', None, u'none'}:
            field_type = u'char'

        if table_id.startswith('#'):
            pass
        elif table_id:
            # start new table
            table_id = table_id.replace('.0', '')
            table = Table(u'rovat_{}'.format(table_id), field_desc)
            tables.append(table)
        elif field_name in ('rovat', 'alrovat', '-'):
            # fields with special meaning
            pass
        elif field_name:
            table.add(Field(field_name, field_desc, field_length, field_type))

    return tables


def _format_field_for_human(field):
    return (
        u'{o.name:12s} {o.description:32}  # {o.type} ({o.length})'
        .format(o=field)
    )


def _format_fields_for_human(fields, indent=u' - '):
    return u'\n'.join(indent + _format_field_for_human(field) for field in fields)


def format_table_for_human(table):
    '''
    Format a table for human consumption
    '''
    header = u'{o.name}    -    {o.description}'.format(o=table)
    header += '\n' + ('-' * max(61, len(header)))
    return (
        u'{header}\n{fields}'
        .format(header=header, fields=_format_fields_for_human(table.fields))
    )


def format_tables_for_human(tables):
    '''
    Format list of tables for human consumption
    '''
    return u'\n\n'.join(format_table_for_human(table) for table in tables)


def print_tables():
    tables = read_tables()
    print(format_tables_for_human(tables).encode('utf-8'))


def main():
    logging.basicConfig()
    print_tables()


if __name__ == '__main__':
    main()
