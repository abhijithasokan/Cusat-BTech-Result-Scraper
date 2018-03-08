import re
from urllib import request, parse
import time
import os
from random import randint
from tqdm import tqdm
from threading import Thread
import sys


class CusatResultScraper():
	'''

	If you have libreoffice calc installed in your system, this script will generate the
	output as spreadsheet.

	Example of data for 7th semester result:
		CusatResultScraper(semester="7",
			class_name = 'CS-B',
			result_month = 11,
			result_year = 2017,
			exam_type = "Regular",
			subject_codes = ['1701', '1702', '1703', '1704', '1705', '17L1', '17L2', '17L3', '17L4']
		).fetchResults()



	Example of data for 3rd semester result:
		CusatResultScraper(semester="3",
			class_name = 'CS-B',
			result_month = 11,
			result_year = 2015,
			exam_type = "Regular",
			subject_codes = ['1301', '1302', '1303', '1304', '1305', '1306', '13L1', '13L2']
		).fetchResults()



	Example of data for 1&2 semester Supplementary result:
		CusatResultScraper(semester="1&2",
			class_name = 'CS-B',
			result_month = 6,
			result_year = 2016,
			exam_type = "Supplementary",
			subject_codes = ['1101', '1102', '1103', '1104', '1105', '1106', '1107', '1108', '1109', '11L1', '11L2', '11L3'],
			log_errors = True,
		).fetchResults()

	'''
	
	RESULT_URL = 'http://exam.cusat.ac.in/erp5/cusat/CUSAT-RESULT/Result_Declaration/display_sup_result'

	ROLL_NUMBERS = {
		'CS-B' :	['12150800', '12150801', '12150803', '12150872', '12150813', '12150877', '12150815', '12150879', '12150881', '12150817', '12150818', '12150819', '12150821', '12150883', '12150822', '12150886', '12150825', '12150826', '12150887', '12150888', '12150889', '12150828', '12150829', '12150830', '12150890', '12150831', '12150832', '12150835', '12150836', '12150837', '12150838', '12150840', '12150842', '12150893', '12150896', '12150843', '12150846', '12150898', '12150899', '12150902', '12150848', '12150855', '12150904', '12150857', '12150858', '12150860', '12150861', '12150907', '12150910', '12150914', '12150915', '12150917', '12150918', '12150919', '12150921', '12150867', '12150869', '12150870', '12150871', '12140884', '12140847', '12150925','12150926', '12150928' , '12150927'  ],
		'CS-A' :    ['12150600', '12150802', '12150804', '12150805', '12150806', '12150807', '12150808', '12150809', '12150810', '12150811', '12150812', '12150814', '12150816', '12150820', '12150823', '12150824', '12150827', '12150833', '12150834', '12150839', '12150841', '12150844', '12150845', '12150847', '12150849', '12150850', '12150851', '12150852', '12150853', '12150854', '12150856', '12150859', '12150862', '12150863', '12150864', '12150866', '12150868', '12150873', '12150874', '12150875', '12150876', '12150880', '12150882', '12150884', '12150885', '12150891', '12150892', '12150894', '12150895', '12150897', '12150900', '12150901', '12150903', '12150905', '12150906', '12150908', '12150909', '12150911', '12150912', '12150913', '12150916', '12150920', '12150922', '12150923', '12150924', '12150929', '12150930'],
		'EB'   :	['10150800', '10150801', '10150802', '10150816', '10150803', '10150817', '10150818', '10150819', '10150820', '10150821', '10150822', '10150823', '10150824', '10150825', '10150826', '10150827', '10150828', '10150829', '10150830', '10150831', '10150832', '10150833', '10150804', '10150805', '10150834', '10150835', '10150806', '10150836', '10150837', '10150838', '10150807', '10150839', '10150840', '10150808', '10150809', '10150810', '10150841', '10150842', '10150811', '10150812', '10150843', '10150813', '10150844', '10150845', '10150814', '10150846', '10150847', '10150848', '10150849', '10150815', '10150850', '10150851', '10150852', '10150857', '10150853', '10150858', '10150854', '10150855', '10150856'	],
		'EC-A' : 	['13150800', '13150806', '13150807', '13150859', '13150809', '13150810', '13150811', '13150813', '13150862', '13150864', '13150865', '13150867', '13150868', '13150870', '13150872', '13150818', '13150873', '13150819', '13150820', '13150874', '13150877', '13150825', '13150879', '13150827', '13150883', '13150884', '13150886', '13150828', '13150888', '13150830', '13150833', '13150889', '13150892', '13150836', '13150837', '13150838', '13150893', '13150894', '13150895', '13150897', '13150898', '13150900', '13150839', '13150902', '13150841', '13150843', '13150904', '13150905', '13150907', '13150908', '13150909', '13150845', '13150846', '13150911', '13150847', '13150915', '13150849', '13150918', '13150920', '13150922', '13150696', '13150924', '13150926', '13150928', '13150930', '13150931', '13150932'	],
		'EC-B' : 	['13150801', '13150802', '13150803', '13150804', '13150805', '13150856', '13150857', '13150858', '13150808', '13150812', '13150860', '13150861', '13150814', '13150816', '13150863', '13150866', '13150869', '13150871', '13150817', '13150821', '13150822', '13150823', '13150824', '13150875', '13150876', '13150878', '13150826', '13150880', '13150881', '13150882', '13150885', '13150887', '13150829', '13150831', '13150832', '13150834', '13150890', '13150835', '13150891', '13150896', '13150899', '13150901', '13150840', '13150842', '13150903', '13150844', '13150906', '13150910', '13150912', '13150913', '13150914', '13150848', '13150916', '13150917', '13150919', '13150850', '13150851', '13150855', '13150921', '13150852', '13150853', '13150854', '13150923', '13150933', '13150925', '13150927', '13150929', '13150934'	],
		'EE'   :	['19150800', '19150801', '19150835', '19150836', '19150802', '19150837', '19150803', '19150804', '19150838', '19150805', '19150806', '19150807', '19150808', '19150809', '19150839', '19150810', '19150811', '19150840', '19150812', '19150841', '19150813', '19150814', '19150815', '19150816', '19150842', '19150817', '19150843', '19150844', '19150818', '19150845', '19150819', '19150820', '19150821', '19150822', '19150823', '19150824', '19150846', '19150825', '19150847', '19150848', '19150849', '19150850', '19150851', '19150826', '19150852', '19150827', '19150828', '19150853', '19150854', '19150855', '19150856', '19150829', '19150830', '19150831', '19150857', '19150858', '19150859', '19150832', '19150833', '19150834', '19150601', '19150671', '19150860', '19150861', '19150862', '19150863', '19150864', '19150865'	]
	}

	GRADES_RE = re.compile('Total.*?:(.*?)<br>.*?GPA.*?:(.*?)<br>',re.DOTALL)
	NAME_RE = re.compile('<th>.*?Student Name.*?</th>.*?<td>(.*?)</td>',re.DOTALL)
	HEADER = {
		'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
		'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
		'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
		'Accept-Encoding': 'none',
		'Accept-Language': 'en-US,en;q=0.8',
		'Connection': 'keep-alive'
	}
	TRY_LIMIT = 10


	def __init__(self,semester,class_name,result_month,result_year,exam_type,subject_codes,log_errors=False):
		
		self.register_numbers = CusatResultScraper.ROLL_NUMBERS[class_name]
		self.subject_codes = subject_codes
		self.branch = class_name.partition('-')[0]
		self.file_name = ( 'Semester_%s_%s_result_%s.txt'%(semester,exam_type,class_name) ).replace('&','and')
		self.log_errors = log_errors
		ind_data = {
			"statuscheck" : "failed",
			"deg_name" : "B.Tech",
			"semester" : semester,
			"month" : ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"][result_month-1],
			"year" : result_year,
			"result_type" : exam_type,    ########
			"ipadress" : "106.208.%d.%d"%(randint(1,254),randint(1,254)),
			"date_time" : time.strftime("%Y/%m/%d")+' {}:{}:{}.350'.format(time.strftime("%H"),time.strftime("%M"),time.strftime("%S"))+' GMT+0530',
		}

	
		self.marks_re = re.compile('<td>(%s.*?)</td>.*?<td>.*?</td>.*?<td.*?>\s*(.*?)(\(\w\))\s*</td>'%(self.branch), re.DOTALL)

		self.write_file = open(self.file_name,'w')

		self.subjectCodeToPos = dict( (self.branch+sub,pos) for pos,sub in enumerate(self.subject_codes) )

		self.post_data_common = parse.urlencode(ind_data).encode() + bytes('&regno=',encoding='utf-8')

		header = 'NAME  , REG NO ,   %s   ,  TM  ,  GPA  \n'%('   ,   '.join(self.subject_codes))
		self.write_file.write(header)

		self.req = request.Request(CusatResultScraper.RESULT_URL, headers=CusatResultScraper.HEADER)

		print("\t\tDETAILS\nSemester: %s\nClass: %s\nExam date: %s %d\n"%(semester,class_name,ind_data['month'],result_year))
		self.pbar = tqdm(total=len(self.register_numbers),mininterval=0.1)



	def fetchSingleResult(self,regno,pos):

		post_data = self.post_data_common + bytes(str(regno),encoding='utf8')

		count = 0
		timeout = 0.5
		reason = ''
		while True:
			if count>=CusatResultScraper.TRY_LIMIT:
				if self.log_errors:
					self.reasons[pos] = reason
				break
			try:
				
				result_page = request.urlopen(self.req,data = post_data, timeout=timeout).read().decode()

				if str(regno) not in result_page:
					reason = "%s not appeared for exam "%regno
					count = CusatResultScraper.TRY_LIMIT
					continue
				_tmarks = self.marks_re.findall(result_page)
				tmarks = []
				for sub_code,mark,grade in _tmarks:
					tmarks.append( (sub_code.replace(' ',''),mark,grade) )

				name = CusatResultScraper.NAME_RE.search(result_page).group(1)
				grades = CusatResultScraper.GRADES_RE.search(result_page)

				gpa = (grades.group(1),grades.group(2))

				marks = [('','','')]*len( self.subject_codes )
				
				for item in tmarks:
					if item[0] not in self.subjectCodeToPos:
						end_pos = item[0].rfind('E')
						item = (item[0][:end_pos],item[1],item[2])
					marks[self.subjectCodeToPos[item[0]]] = item
				
				marks_toprint = ' , '.join( '%3s%3s'%(x[1],x[2]) for x in marks )

				gpa_toprint = '  %s  ,  %s  '%(gpa[0],gpa[1])

				final_details = '%s , %s   , %s ,%s'%(name, regno, marks_toprint ,gpa_toprint)
					
				#print(final_details)
				self.results[pos] =  final_details

				break

			except (request.URLError,OSError,ConnectionResetError) as e:
				count+=1
				timeout += 0.25
				reason = str(e)
				continue
			except KeyboardInterrupt:
				print('Exit by KeyboardInterrupt')
				exit()
			except Exception as e:
				count = CusatResultScraper.TRY_LIMIT
				reason = str(e)

		self.pbar.update(1)



	def fetchResults(self):
		self.reasons = ['']*len(self.register_numbers)
		self.results = ['']*len(self.register_numbers)
		threads = []
		for pos,regno in enumerate(self.register_numbers):
			t = Thread(target=self.fetchSingleResult, args=(regno,pos))
			t.start()
			threads.append(t)
		
		for t in threads:
			t.join()

		self.pbar.refresh()

		if self.log_errors:
			print(*self.reasons,sep='\n')

		self.write_file.write('\n'.join(self.results))	

		print('\n\n\nNo.of results fetched: %d'% (len(self.results) - self.results.count('')) )
		self.write_file.close()
		os.system("libreoffice --calc --convert-to xlsx %s"%self.file_name)
		os.system("rm %s"%self.file_name)



if __name__ == '__main__':
	if '--help' in sys.argv:
		print(CusatResultScraper.__doc__)
	else:
		CusatResultScraper(semester="7",
			class_name = 'CS-B',
			result_month = 11,
			result_year = 2017,
			exam_type = "Regular",
			subject_codes = ['1701', '1702', '1703', '1704', '1705', '17L1', '17L2', '17L3', '17L4'],
			log_errors = False
		).fetchResults()


