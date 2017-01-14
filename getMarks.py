'''
============== Documentation ===============

Enter the details about the examination in "data_of_exam" (line 66) dictionary given below
The result will be printed to "result.txt" file in the present working directory.

The roll numbers used here are those of MEC students.
If you're non MECian provide the required roll numbers as a list
in variable : "register_numbers" (line 86).

If you have libreoffice calc installed in your system, this script will generate the
output as spreadsheet.


Example of data for 3rd semester result:

		data_of_exam = {
			'month' : "November",		 // Month of examination
			'year' : "2015",			 // Year
			'semester' : "3",			 // Semester
			'type_of_exam' : "Regular",  // Regular or supply
			'run_time' : "2017/01/12",	 // Date of running the script
			'class' : 'CS-B',			 // Class, it can be either "CS-A" or "CS-B",
			'branch' : 'CS',
			'subjectCodes' : ['CS1301', 'CS1302', 'CS1303', 'CS1304', 'CS1305', 'CS1306', 'CS13L1', 'CS13L2'] // The Subject Codes You want 
		}

Example of data for 1&2 revalution result:

		data_of_exam = {
		'month' : "May",
		'year' : "2015",
		'semester' : "1&2",
		'type_of_exam' : "Revaluation",
		'run_time' : "2017/01/12",
		'class' : 'CS-B',
		'branch' : 'CS',
		'subjectCodes' : ['CS1101', 'CS1102', 'CS1103', 'CS1104', 'CS1105', 'CS1106', 'CS1107', 'CS1108', 'CS1109', 'CS11L1', 'CS11L2', 'CS11L3']

	}


Example of data for 4th semester result:

		data_of_exam = {
			'month' : "April",
			'year' : "2016",
			'semester' : "4",
			'type_of_exam' : "Regular",
			'run_time' : "2017/01/12",
			'class' : 'CS-B',
			'branch' : 'CS',
			'subjectCodes' : ['CS1401', 'CS1402', 'CS1403', 'CS1404', 'CS1405', 'CS1406', 'CS14L1', 'CS14L2']
		}

'''

import re
import urllib2
import urllib
import time
import os



data_of_exam = {
	'month' : "May",
	'year' : "2015",
	'semester' : "1&2",
	'type_of_exam' : "Revaluation",
	'run_time' : "2017/01/12",
	'class' : 'CS-B',
	'branch' : 'CS',
	'subjectCodes' : ['CS1101', 'CS1102', 'CS1103', 'CS1104', 'CS1105', 'CS1106', 'CS1107', 'CS1108', 'CS1109', 'CS11L1', 'CS11L2', 'CS11L3']
	}


file_name = 'Semester_%s_%s_result_%s.txt'%(data_of_exam['semester'],data_of_exam['type_of_exam'],data_of_exam['class'])
file_name =file_name.replace('&','and')
roll_nos = {
	'CS-B' :	['12150800', '12150801', '12150803', '12150872', '12150813', '12150877', '12150815', '12150879', '12150881', '12150817', '12150818', '12150819', '12150821', '12150883', '12150822', '12150886', '12150825', '12150826', '12150887', '12150888', '12150889', '12150828', '12150829', '12150830', '12150890', '12150831', '12150832', '12150835', '12150836', '12150837', '12150838', '12150840', '12150842', '12150893', '12150896', '12150843', '12150846', '12150898', '12150899', '12150902', '12150848', '12150855', '12150904', '12150857', '12150858', '12150860', '12150861', '12150907', '12150910', '12150914', '12150915', '12150917', '12150918', '12150919', '12150865', '12150921', '12150867', '12150869', '12150870', '12150871', '12140884', '12140847', '12150925','12150926', '12150928' , '12150927'  ],
	'CS-A' :    ['12150600', '12150802', '12150804', '12150805', '12150806', '12150807', '12150808', '12150809', '12150810', '12150811', '12150812', '12150814', '12150816', '12150820', '12150823', '12150824', '12150827', '12150833', '12150834', '12150839', '12150841', '12150844', '12150845', '12150847', '12150849', '12150850', '12150851', '12150852', '12150853', '12150854', '12150856', '12150859', '12150862', '12150863', '12150864', '12150866', '12150868', '12150873', '12150874', '12150875', '12150876', '12150880', '12150882', '12150884', '12150885', '12150891', '12150892', '12150894', '12150895', '12150897', '12150900', '12150901', '12150903', '12150905', '12150906', '12150908', '12150909', '12150911', '12150912', '12150913', '12150916', '12150920', '12150922', '12150923', '12150924', '12150929', '12150930']
}	


register_numbers = roll_nos[data_of_exam['class']]


header = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
       'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
       'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
       'Accept-Encoding': 'none',
       'Accept-Language': 'en-US,en;q=0.8',
       'Connection': 'keep-alive'}




marks_str = '''<td>.*?</td><td style="text-align:center;">

(.*?)\((\w)\)

</td>'''

marks_str = '''<td>(%s.*?)</td><td>.*?</td><td style="text-align:center;">

(.*?)(\(\w\))

</td>'''%(data_of_exam['branch'])

grades_str = '''</table>

                  Total :  (.*?)<br>
         GPA   :  (.*?)<br>'''


f = open(file_name,'w')


noOfSubs = len(data_of_exam['subjectCodes'])
subjectCodeToPos = dict(t[::-1] for t in enumerate(data_of_exam['subjectCodes']))



count_data = 0
subs = '   ,   '.join(data_of_exam['subjectCodes']) 
head = 'NAME  ,'+' REG NO ,   '+subs+'   ,  TM  ,  GPA  '

f.write(head+'\n')


for regno in register_numbers:
	
	count = 0
	while True:
		if count==1:
			print 'Breaked due to either timeout or %s not appearing for exam'%regno
			break
		try:
			t = data_of_exam['run_time']+' {}:{}:{}.350'.format(time.strftime("%H"),time.strftime("%M"),time.strftime("%S"))+' GMT+0530'

			ind_data = {
				"regno":regno,
				"statuscheck":"failed",
				"deg_name":"B.Tech",
				"semester": data_of_exam['semester'],
				"month":data_of_exam['month'],
				"year":data_of_exam['year'],
				"result_type": data_of_exam['type_of_exam'],
				"ipadress":"117.245.10.120",
				"date_time":t,
			}
			next_url = 'http://exam.cusat.ac.in/erp5/cusat/CUSAT-RESULT/Result_Declaration/display_sup_result'


			ind_data = urllib.urlencode(ind_data)
			req = urllib2.Request(next_url,ind_data,headers=header)
			res_page = urllib2.urlopen(req).read()

			

			tmarks = re.findall(marks_str,res_page)
			

			name = re.search('<th>Student Name</th><td>(.*?)</td>',res_page).group(1)
			grades = re.search(grades_str,res_page)

			try:
				gpa = (grades.group(1),grades.group(2))
			except:
				gpa = ('   ','   ')

			marks = [('','','')]*noOfSubs
			for item in tmarks:
				marks[subjectCodeToPos[item[0]]] = item
			
			marks_toprint = ' , '.join(map(lambda x:'%3s%3s'%(x[1],x[2]),marks) )

			name_toprint = name

			gpa_toprint = '  %s  ,  %s  '%(gpa[0],gpa[1])

			final_details = name_toprint+' , '+regno+'   , '+marks_toprint+' ,'+gpa_toprint
			print final_details
			f.write(final_details+'\n')


			count_data+=1
			break

		except:
			count+=1
			continue
	



print '\n\n\n\n\n\nNo.of results fetched: %d'%count_data
f.close()
os.system("libreoffice --calc --convert-to xlsx %s"%file_name)
os.system("rm %s"%file_name)


