import configparser
from FilledForm import *



if __name__=='__main__':
    filled_from_id_list=[66]
    config = configparser.ConfigParser()
    config.read('config.ini')
    database_info=config['DATABASE']
    table_name=config['TABLE_NAME']

    fmt = "%(asctime)-15s %(levelname)s %(filename)s %(lineno)d %(process)d %(message)s"
    datefmt = "%a %d %b %Y %H:%M:%S"
    logging.basicConfig(filename='logger.log', level=logging.INFO,format=fmt,datefmt=datefmt)



    filled_form=FilledForm(filled_from_id_list[0])
    filled_form.init_form(database_info,table_name)

    for from_id in filled_from_id_list:
        filled_form.add_form(from_id,database_info,table_name)
    filled_form.save_form('example.xls')



