import configparser,sys
from FilledForm import *



if __name__=='__main__':
    config = configparser.ConfigParser()
    config.read('config.ini')
    database_info=config['DATABASE']
    table_name=config['TABLE_NAME']

    fmt = "%(asctime)-15s %(levelname)s %(filename)s %(lineno)d %(process)d %(message)s"
    datefmt = "%a %d %b %Y %H:%M:%S"
    logging.basicConfig(filename='logger.log', level=logging.INFO,format=fmt,datefmt=datefmt)



    if len(sys.argv)<2:
        print('Error argv missing form id list')
        logging.info('Error argv missing form id list')
    else:
        filled_arg=sys.argv[1]
        export_arg=sys.argv[2]

        filled_from_id_list=filled_arg.split('=')[1].split(',')
        export_path=export_arg.split('=')[1]

        filled_from_id_list=[index for index in range(85,193)]

        try:
            filled_form=FilledForm(filled_from_id_list[0])
            filled_form.init_form(database_info,table_name)
            if filled_form==None:
                print("Failed to init form")
                logging.info("Failed to init form")
            else:
                for from_id in filled_from_id_list:
                    filled_form.add_form(from_id,database_info,table_name)
                filled_form.save_form(export_path)
                print("Success export form to {}".format(export_path))
                logging.info("Success export form to {}".format(export_path))
        except:
            print("Failed to export form")
            logging.info("Failed to export form")




