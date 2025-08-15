/*	PerfusionPDF
 *	Import perfusion record PDF reports
 */

#Requires AutoHotkey v2

fname := ".\data\Record 6_23_2022.pdf"
y := record(fname)

/*	====================================================================================
 */
Class record
{
	__New(fileIn) {
		if FileExist(fileIn) {
			SplitPath(fileIn,&fName,&path,&ext,&fNam)
			this.file := []
			this.file.path := path
			this.file.fName := fName
			this.file.fNam := fNam
			this.file.ext := ext
			this.file.fOut := A_Now ".txt"
			this.exe := ".\includes\pdftotext.exe"
			this.opts := "-table -eol unix -nopgbrk -margint 108 -marginb 90"

			this.readFile(fileIn)

			return this
		} else {
			return Error
		}
	}

	readFile(fileIn) {
		if (this.file.ext="txt") {
			this.text := FileRead(fileIn)
			return
		}
		RunWait(this.exe " " this.opts " `"" this.file.fName "`" " this.file.fOut , this.file.path,"Hide")
		txtIn := FileRead(this.file.path "\" this.file.fOut)
		txtIn := StrReplace(txtIn,"`n`n","`n")
		FileAppend(txtIn,this.file.path "\" A_Now "-out.txt")
		this.text := txtIn
		; FileDelete(this.file.fOut)
	}
}

cleanspace(&txt) {
	txt := StrReplace(txt,"`n"," ")
	txt := StrReplace(txt," . ",". ")
	loop 
	{
		txt := StrReplace(txt,"  "," ",, &count)
		if (count=0)	
			break
	}
}

#Include includes
#include strx2.ahk
