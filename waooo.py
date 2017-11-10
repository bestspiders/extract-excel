#-*- coding:utf-8 -*-
import xlwings as xw
import yaml
import re,os
print os.getcwd()

# source=yaml_content['source']
class normal_start():
	def __init__(self,sheetname,match_name,source,nameposition,dataposition,deep,number,rename):
		self.rename=rename
		self.match_name=match_name
		self.sheetname=sheetname
		self.source=source
		self.deep=deep
		self.nameposition=nameposition
		self.dataposition=dataposition
	def value_content(self):
		sht=wb.sheets[self.sheetname]
		nas_namelist=sht.range('A1:A'+str(self.deep)).value
		nas_valuelist=sht.range(self.source+'1:'+self.source+str(self.deep)).value
		for k in range(0,len(nas_namelist)):
			if nas_namelist[k]==match_name:
				if number>0:
					return round(float(nas_valuelist[k]),number)
				else:
					return 	int(float(nas_valuelist[k]))
	def run(self):
		sht=wb.sheets[self.sheetname]
		nas_namelist=sht.range('A1:A'+str(self.deep)).value
		nas_valuelist=sht.range(self.source+'1:'+self.source+str(self.deep)).value
		for k in range(0,len(nas_namelist)):
			if nas_namelist[k]==match_name:
				if number>0:
					will_sht.range(dataposition).value=str(round(float(nas_valuelist[k]),number))
				else:
					will_sht.range(dataposition).value="'"+str(int(float(nas_valuelist[k])))
				if rename:
					will_sht.range(nameposition).value=rename
					break
				else:
					will_sht.range(nameposition).value=match_name
					break
		will_sht.range(nameposition).value=match_name


if __name__ == '__main__':
	half_path=os.getcwd().split('\\')
	new_half_path='\\\\'.join(half_path)
	all_file=os.listdir('.')
	now_file=[new_half_path+'\\\\'+every_path for every_path in all_file]
	write_file=new_half_path+'\\\\'+u'产品及系统运行质量分析周报.xlsx'
	will_write=xw.Book(write_file)#打开写入的excel
	will_sht=will_write.sheets[0]
	all_yaml_file=open('.\\mkyaml\\woo.txt')
	all_content=all_yaml_file.read()
	all_yaml_file.close()
	many_path=all_content.split('\n')
	for every_yaml in many_path:
		file=open('.\\mkyaml\\'+every_yaml)
		yaml_content=yaml.load(file)#载入yaml
		file_match=yaml_content['name']
		for every_file in now_file:
			if re.search(file_match,every_file):
				file_name=every_file
		wb = xw.Book(file_name)#打开要读取的excel
		all_task=yaml_content['content']
		global a_d
		a_d={}
		for n in range(0,len(all_task)):
			deep=300
			number=1
			sheetname=''
			nameposition,dataposition=1,1
			rename=None
			read_dic=all_task[n]['task'+str(n+1)]
			if  read_dic.has_key('sheet'):
				sheetname=read_dic['sheet']
			if read_dic.has_key('name'):
				match_name=read_dic['name']
			if read_dic.has_key('deep'):
				deep=read_dic['deep']
			if read_dic.has_key('decimal'):
				number=int(read_dic['decimal'])
			if read_dic.has_key('nameposition'):
				nameposition=read_dic['nameposition']
			if read_dic.has_key('dataposition'):
				dataposition=read_dic['dataposition']
			if read_dic.has_key('source'):
				source=read_dic['source']
			if read_dic.has_key('rename'):
				rename=read_dic['rename']
			if nameposition==1 or dataposition==1:
				a_d['task'+str(n+1)]=normal_start(sheetname,match_name,source,nameposition,dataposition,deep,number,rename).value_content()
			elif read_dic.has_key('calculate'):
				calculate=read_dic['calculate']
				calculate=re.sub('"','',calculate)
				get_result=eval(calculate)
				will_sht.range(nameposition).value=match_name
				will_sht.range(dataposition).value=get_result
				a_d={}
			else:
				normal_start(sheetname,match_name,source,nameposition,dataposition,deep,number,rename).run()