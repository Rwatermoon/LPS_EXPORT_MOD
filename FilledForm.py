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

    def decode_division(self,option_code):
        division_code_source=defaultdict(str)
        with open('meta_division.txt','r',encoding='utf8') as meta_file:
            for line in meta_file:
                code_list=line.split('\t')
                division_code_source[code_list[0]]=code_list[1]

        result=""
        province=division_code_source[option_code]
        city=division_code_source[option_code[:-2]+'00']
        county=division_code_source[option_code[:-4]+'0000']
        result=province+'/'+city+'/'+county


        return result
    def init_form(self,database_info,table_name):
        try:
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
        except AttributeError as e:
            print(e)
            logging.info("{} in init_form".format(e))
        finally:
            conn.close()

    def add_form(self,filled_from_id_list,database_info,table_name):
        try:
            conn = mysql.connector.connect(host=database_info['host'], user=database_info['user'],
                                           password=database_info['password'],
                                           database=database_info['database'], use_unicode=True)
            cursor = conn.cursor()
            sql_str = 'select * from {} where filled_id in {}'.format(table_name['filled_question'], tuple(filled_from_id_list))
            print(sql_str)
            cursor.execute(sql_str)
            filled_questions = cursor.fetchall()
            for q in filled_questions:
                self.filled_question_qid[q[3]].append(q)

            # logging.info("Success read form with ID={}".format(form_id))
        except AttributeError as e:
            print(e)
            logging.info("{} in add_form".format(e))
        finally:
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
                if sec_question['type']=='division':
                    for filled_q in self.filled_question_qid[sec_question['qid']]:
                        option_code = filled_q[4]
                        option_value=self.decode_division(str(option_code))
                        ws.write(row_num, line_num, option_value)
                        row_num += 1
                    line_num = line_num + 1

                if sec_question['type']=='table':
                    table_content=self.filled_question_qid[sec_question['qid']]
                    if len(table_content)>0:
                        table_content_str=table_content[0][5]
                        table_content_row_list=table_content_str[2:-3].split('],[')


                        for row_index in range(len(table_content_row_list)):
                            table_content_line_list=table_content_row_list[row_index].split(',')
                            for line_index in range(len(table_content_line_list)):
                                ws.write(row_num+row_index,line_num+line_index,table_content_line_list[line_index])
                    else:
                        for table_item in sec_question['options']:
                            print(table_item)
                            ws.write(row_num,line_num,table_item)
                            line_num+=1
                        line_num=0
                        row_num+=1
                        for table_item in sec_question['extras']:
                            print(table_item)
                            ws.write(row_num,line_num,table_item)
                elif sec_question['type']=='single':
                    for filled_q in self.filled_question_qid[sec_question['qid']]:
                        option_index=filled_q[4]-1
                        if option_index<0:option_index=0
                        if option_index>len(sec_question['options'])-1:option_index=len(sec_question['options'])-1
                        option_val=sec_question['options'][option_index]
                        ws.write(row_num, line_num, option_val)
                        row_num += 1
                    line_num=line_num+1
                elif sec_question['type']=='multi':
                    for filled_q in self.filled_question_qid[sec_question['qid']]:
                        option_val=""
                        for bin_index in range(len(sec_question['options'])):
                            if (filled_q[4]&(2**bin_index))!=0:
                                option_val+=sec_question['options'][bin_index]+','
                        ws.write(row_num, line_num, option_val)
                        row_num += 1
                    line_num=line_num+1

                    pass
                else:
                    for filled_q in self.filled_question_qid[sec_question['qid']]:
                        ws.write(row_num, line_num,filled_q[5])
                        row_num += 1
                    line_num=line_num+1
        wb.save(save_path)
        logging.info("Success export to {}".format(save_path))
        # except IndexError as e:
        #     print(e)
        #     logging.info("{} in save_form".format(e))


