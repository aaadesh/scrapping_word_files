from docx2python import docx2python
import pandas as pd
import os

folder = r"H:\Word Files"
files = os.listdir(folder)
#print(files)

# extract docx content

df1 = pd.DataFrame()

def get_file_data(file):

	data = docx2python(file)

	#for a in range(len(data.body)):
	#	print(a)
	#	print(data.body[a])
	x = data.header[3][0] + data.header[4][0]
	print(data.header[3][0])
	print(data.header[4][0])
	print(x)

	collection = [file, data.header[3][0], data.body[10], data.body[15][0][1], data.body[15][1][1], data.body[15][2][1], data.body[15][3][1], data.body[17][0][1], data.body[17][1][1], data.body[17][2][1], data.body[17][3][1], data.body[17][4][1], data.body[18][0]]
	#print(collection)
	return collection
	

for b in range(len(files)):
	row = get_file_data(folder + "\\" + files[b])
	print("fetching data from:", b, files[b])
	df1[b] = pd.DataFrame(row)


print(df1)

df = df1.transpose()

df.columns = ['Filename', 'Header', 'Analyst/Reviewer', 'Family ID', 'Patent Number', 'Priority', 'Expiry', 'Litigation Grade SEP', 'Mapped Standard', 'Relevant TS/TR', 'Relevant Sections', 'Overlap Percentage', 'Search Observations']


print(df)

output = r"H:\Review\extracted_data.csv"
df.to_csv(output, encoding='utf-8')