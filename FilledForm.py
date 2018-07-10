from collections import defaultdict
import mysql.connector
import json
import xlwt
import logging





class FilledForm:
    def __init__(self, form_id=''):
        self.ini_form_id=form_id
        self.section=[]
        self.question_title=[]
        self.filled_question_qid=defaultdict(list)

    def init_form(self,database_info,table_name):
        conn = mysql.connector.connect(host=database_info['host'], user=database_info['user'],
                                       password=database_info['password'],
                                       database=database_info['database'], use_unicode=True)
        cursor = conn.cursor()

        logging.info("Success connect to {}".format(database_info['host']))
        # Query form_id
        sql_str = 'select form_id from {} where id={}'.format(table_name['filled_form'], self.ini_form_id)

        cursor.execute(sql_str)
        template_from_id = cursor.fetchone()
        if template_from_id is not None:
            template_from_id = template_from_id[0]
        else:
            return None

        # Query form_info
        sql_str = 'select content from {} where id={}'.format(table_name['form_info'], template_from_id)
        cursor.execute(sql_str)
        from_content = cursor.fetchone()
        if from_content is not None:
            from_content = from_content[0]
        else:
            return None


        content_json = json.loads(from_content)
        self.sections = content_json['sections']
        self.questions = content_json['questions']



        for sec_q in self.questions:
            self.filled_question_qid[sec_q['qid']]=[]

        logging.debug("Success load all data")
        conn.close()

    def add_form(self,form_id,database_info,table_name):
        conn = mysql.connector.connect(host=database_info['host'], user=database_info['user'],
                                       password=database_info['password'],
                                       database=database_info['database'], use_unicode=True)
        cursor = conn.cursor()
        sql_str = 'select * from {} where filled_id={}'.format(table_name['filled_question'], form_id)
        cursor.execute(sql_str)
        filled_questions = cursor.fetchall()
        for q in filled_questions:
            self.filled_question_qid[q[3]].append(q)

        logging.info("Success read form with ID={}".format(form_id))
        conn.close()

    def save_form(self,save_path):

        wb = xlwt.Workbook(encoding='utf-8')
        for section in self.sections:
            ws = wb.add_sheet(section['title'], cell_overwrite_ok=True)
            sec_questions = [item for item in self.questions if item['qid'] // 100 == section['sid']]

            line_num = 0
            for sec_question in sec_questions:
                row_num = 0

                ws.write(row_num,line_num,sec_question['title'])
                row_num+=1
                if sec_question['type']=='table':
                    print(sec_question['type'])
                    for table_item in sec_question['extras']:
                        print(table_item)
                        ws.write(row_num,line_num,table_item)
                        row_num+=1
                    for table_item in sec_question['option']:
                        print(table_item)
                        ws.write(row_num,line_num,table_item)
                        row_num+=1
                row_num+=1
                for filled_q in self.filled_question_qid[sec_question['qid']]:
                    ws.write(row_num,line_num,filled_q[4])
                    ws.write(row_num, line_num+1,filled_q[5])
                    row_num += 1
                line_num=line_num+2
        wb.save(save_path)
        logging.info("Success export to {}".format(save_path))
