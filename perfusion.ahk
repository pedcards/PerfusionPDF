/*	PerfusionPDF
 *	Import perfusion record PDF reports
 */

#Requires AutoHotkey v2

fname := ".\data\Record 6_23_2022.pdf"
fname := ".\data\20250814192513-out.txt"
y := record(fname)
n := 0
demog := y.block("\R*Patient Data\R",1,0,"\R+Surgery Team\R",1)
team := y.block("\R+Surgery Team",1,1,"\R+Disposables\R",1)
onlinedata := y.block("\R+Online\s+Data\R",1,0,"\R+Cardioplegia\s+Values\R",1)

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

	/*	Get a block using stRegX
	 *		BS = begin string
	 *		BO = position from start of haystack
	 *		BT = trim BS, true or false
	 *		ES = end string
	 *		ET = trim ES, true or false
	 *		N  = var for next offset
	 */
	block(BS:="",BO:=1,BT:=0,ES:="",ET:=0,&N:=0) {
		res := stRegX(this.text,BS,BO,BT,ES,ET,&N)
		return res
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
#Include strx2.ahk
#Include Peep.v2.ahk