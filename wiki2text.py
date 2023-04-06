import wikipedia
import docx
from tqdm import tqdm


def wiki2text(domainTXTFile):
    # Read the text file to find to content
    with open(domainTXTFile, 'r') as file:
        file_contents = [line.strip() for line in file]

    # Loop the file content for wikipedia title
    for idx,name in tqdm(enumerate(file_contents), total=len(file_contents)):
        page_title = name + " in Malaysia"
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


if __name__ == '__main__':
    domainFile = "sector_malaysia.txt"
    titleTxtFile,urlTxtFile = wiki2text(domainTXTFile=domainFile)
    print()
    print()
    convert2doc(title=titleTxtFile, link=urlTxtFile)
    print("Done")
    



