import wikipedia
import docx
from tqdm import tqdm
import argparse

def wiki2text(domainTXTFile, location):
    # Read the text file to find to content
    with open(domainTXTFile, 'r') as file:
        file_contents = [line.strip() for line in file]

    # Loop the file content for wikipedia title
    for idx,name in tqdm(enumerate(file_contents), total=len(file_contents)):
        page_title = f"{name} in {location}"
        titleName = "Title.txt"
        urlName = "urlLink.txt"
        # try exception error to prevent error occur
        try:
            page = wikipedia.page(page_title)
            url_link = page.url
            # content = re.sub(r'{\|.*?\|}', '', page.content)
            with open(f"{str(idx)}_{page_title}.txt", "w", encoding="utf-8")as f:
                f.write(page.content)
            
            with open(titleName, 'a', encoding="utf-8") as f:
                f.write(f"{url_link}\n")
            
            with open(urlName, 'a', encoding="utf-8") as f:
                f.write(f"{[page.title]}\n")
        except:
            print(f"\nError occur!!! : {page.title}")
    
    return titleName,urlName

def convert2doc(title,link):
    #-------------------------------------------------------
    # convert from text file into word
    with open(title, 'r') as f:
        data = f.readlines()

    with open(link, 'r') as f:
        data2 = f.readlines()

    # Remove newline characters from data
    data = [d.strip() for d in data]
    data2 = [d.strip() for d in data2]

    # create word document
    doc = docx.Document()

    # Add table to document
    table = doc.add_table(rows=len(data), cols=3)
    x=26
    for idx, (d1,d2) in enumerate(zip(data,data2)):
        cell = table.cell(idx,0)
        cell.text = str(idx+x)

        cell1 = table.cell(idx,1)
        cell1.text = d1

        cell2 = table.cell(idx,2)
        cell2.text = d2
    # Save Word document
    doc.save('output2.docx')


if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument("--txt", help="text file for find the information in a list")
    parser.add_argument("--loc", help="location to find")
    args = parser.parse_args()
    if args.txt.endswith(".txt"):
        titleTxtFile,urlTxtFile = wiki2text(domainTXTFile=args.txt,location=args.loc)
        convert2doc(title=titleTxtFile, link=urlTxtFile)
        print("Done")
    else:
        print("It's not a text file")
    



