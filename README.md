## Value Extractor

This is a very simple Java application I've written for a friend who works at a hospital and calibrates and tests devices.

![preview](/images/preview.png)

What this application does is basically that to extract values under the "Value" column from a text file (you can find a sample in [test][1] folder) which is created by [Metron QA-90][2] and write them to an excel (`.xls`) file.

**This application works only with those specifically formatted text files.**

## What's behind 
As you can see from the sample text file, it reads the file line by line until it finds a line which contains the word "value". 

![sample text file](images/values.png)

After the line is found, a flag called `containsValue` is set to `true` to initiate splitting the lines into an array by `;` (semicolons) which the text file has five of them per line after the line which has the word "value".

```java
valuesList.add(line.split(";")[3].trim());
```

To get the values under the "Value" column we only need the element at the 3rd index (4th element) of the array. 

```java
if(containsValue) {
	if(line.isEmpty()) {
	    containsValue = false;
        break;
    }
    valuesList.add(line.split(";")[3].trim());
}
```

Program keeps adding those values into an `ArrayList` called `valuesList` until a blank line is hit which is controlled with `isEmpty()`. We also trim each array element to get rid of the white space at both ends.

Finally we add this value into an `ArrayList` called `valuesList` as you can see above.

## Writing Extracted Values to an Excel File
After adding values into an `ArrayList` we write the values to an excel file with the function below using [Apache POI Library][3].

Generated `.xls` files will be in the same path as their sources.

```java
private void write_toExcel(ArrayList arrayList, File fileName) {
        
	HSSFWorkbook workbook = new HSSFWorkbook();
    HSSFSheet sheet = workbook.createSheet("First Sheet");
    for(int i = 0; i < arrayList.size(); i++) {
	    HSSFRow row = sheet.createRow(i);
        HSSFCell cell = row.createCell(0);
		cell.setCellValue(arrayList.get(i).toString());
    }  
    try{
	    workbook.write(new FileOutputStream(fileName.getCanonicalPath().replace(".txt", "-extracted.xls")));
        workbook.close();
    }catch(IOException e){
	    e.printStackTrace();
    }
}    
```

[1]: https://github.com/andreyuhai/Value-Extractor/tree/master/test
[2]: https://www.celyontecnica.es/var/celyon-1051-qa90MKII.pdf
[3]: https://poi.apache.org/

