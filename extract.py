''' Created by Jeremy Reynolds for UALR COSMOS Team

	Extracts Retweet and mention handles from raw TAGS data.
	Expects two columns [username, tweet] in a .xlsx (headers ignored)
	Outputs four columns [source, target, communicatoinType, weight] to .csv
'''

import pandas as pd, openpyxl, csv

book = openpyxl.load_workbook('Raw Tweets.xlsx')
sheet = book.active

for row in sheet.iter_rows():	# Loops through rows and creates/resets arrays
	user = []					# Contains username handle from row[0]
	mention = []				# Contains mentions with @ as leading character
	retweet1 = []				# Contains retweets with 'RT @' as leading characters
	retweet2 = []				# Contains retweets with 'RT ' as leading characters
	for cell in row:			# Loops through cells within row
		if cell == row[0]:		# If first cell (username), append to user[] array
			user.append(cell.value)
			charCount = 0		# Counter used for retweets with 'RT ' as leading characters
			charR = False		# T/F flag for character R in 'RT '
			charT = False		# T/F flag for character T in 'RT '
			charSpace = False	# T/F flag for the space in 'RT '
			asterisk = False	# T/F flag for retweet/mentions
			rtSpace = False		# T/F flag for retweets without an @ identifier
		else:
			for item in cell.value:
				charCount += 1						# Increases charCount if retweet with no '@'
				if item == ':' and asterisk:		# If colon and astrisk is True, append mention[] to retweet1[]
					retweet1.append(mention)
					mention = []					# Clears mention[]
				if item == ' ' or item == ':' or item == '\n':
					asterisk = False				# If space, colon, or newline, done looking for RT/Mention, flag False
					rtSpace = False
				if item == '@':						# If '@' asterisk (mention) is True
					asterisk = True
				if asterisk == True:				# If asterisk is True, append to mention[]
					mention.append(item)
				if charCount == 1 and item == 'R':	# If first char is 'R', set charR True
					charR = True
				if charCount == 2 and item == 'T' and charR:	# If first two chars are 'RT',set charT True
					charT = True
				if charCount == 3 and item == ' ' and charT:	# If first three chars are 'RT ',set charSpace True
					charSpace = True
				if charCount == 4 and item != '@' and charSpace:# If fourth char is not @ and charSpace is True, set rtSpace True
					rtSpace = True
				if rtSpace:
					retweet2.append(item)			# If rtSpace is True, append item to retweet2[]
					retweet1.append(retweet2)		# Then append retweet2[] to retweet1[] such that array is [['items']]
					
	if retweet1:				# If retweet1[] is not empty join and split by '@'
		handles = ''.join(retweet1[0]).split('@')
		if handles[0] == '':	# If first item is empty, pop()
			handles.pop(0)
		for item in handles:	# Saves username, @handle, retweet to .csv
			with open('output.csv', 'a', newline='') as file:
				writer = csv.writer(file)
				writer.writerow([user[0], item, 'retweet'])
			
	if mention:					# If mention[] is not empty join and split by '@'
		handles = ''.join(mention).split('@')
		handles.pop(0)
		if handles[0] == '':	# If first item is empty, pop()
			handles.pop(0)
		for item in handles:	# Saves username, @handle, mention to .csv
			with open('output.csv', 'a', newline='') as file:
				writer = csv.writer(file)
				writer.writerow([user[0], item, 'mention'])

# Creates dataframe from 'output.csv'
df = pd.read_csv('output.csv', header=None)

# Counts and drops duplicates. Creates new column 'weight' with duplicate sums
df.columns = ['source', 'target', 'communicationType']
df = df.groupby(['source', 'target', 'communicationType']).size().reset_index()
df.rename(columns = {0: 'weight'}, inplace=True)

# Saves to output file, comment out if dupicates should be kept
df.to_csv('output.csv', index=False, encoding='utf-8')
