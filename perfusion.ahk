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
ta := y.getOnlineData(onlinedata)
y.outputCSV(ta)

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

	getOnlineData(txt) {
	/*	Parse txt block for Online Data fields
	 *	After header, scan for column labels and cell coords
	 *	Read each cell
	 */
		/*	First put all lines into an array
		*/
		txtarray := []
		loop parse txt, "`n" {
			txtarray.Push(A_LoopField)
		}

		/*	Read each line in sequence
		 *	"Online Data" triggers column rescan
		 *	"Time" should always be first column
		 *	Look ahead for time-contaning line,
		 *	may need to concat time + AM/PM lines
		 *	Store as timearray, objects indexed to timestamp	
		 */
		timearray := Map()
		while (A_Index<=txtarray.length) 
		{
			txtline := txtarray[A_Index]
			if (txtline="") {
				continue
			}
			if (txtline~="Online\s+Data") {												; Line is Online Data
				cols := getHeaders(A_Index)
				A_Index := A_Index + cols.height
				continue
			}
			try {
				row := getCels(A_Index,cols.col)
				A_Index := A_Index + row.height
				dt := row.cell[1]
				try timearray[dt]
				catch
				{
					timearray[dt] := Map()
				}
				for k,val in cols.hdr {
					if (k=1) {
						continue
					}
					timearray[dt][val] := row.cell[k]
				}
			}
		}
		return timearray

		getHeaders(idx) {
		/*	Get column names and X coords
		*/
			hdrIdx := idx+1																; Header is first line after "Online Data"
			line := txtarray[hdrIdx]
			line := RegExReplace(line,"  P  "," P   ")									; Adjust P in headers
			colx := findHeaders(line)													; Get positions of first row of header
			
			header := Map()
			loop 6 {																	; Look ahead next 6 lines
				row := A_Index
				x := readRow(txtarray[idx+row],colx)
				if (x[1]~="^\d{1,2}:\d{2}") {											; Until finds time string in first cell
					for key in header {
						tx := header[key]
						cleanspace(&tx)
						tx := RegExReplace(tx,"(\[.*?\])")
						header[key] := Trim(tx)
					}
					return {hdr:header,col:colx,height:row-1}
				}
				for key,val in x {														; Concatenate lines of header
					tx := x[key]
					if (row=1) {
						header[key] := tx
					} else {
						header[key] .= tx
					}
				}
			}
		}

		findHeaders(txt) {
		/*	Match position of  first char after "\s\s"
		*/
			n := 0
			res := []
			loop {
				n := RegExMatch(txt,"^\S|(?<=\s\s)\S",&match,n+1)
				if !n {
					break
				}
				res.Push(match.pos)
			}
			return res
		}

		readRow(txt,colx) {
		/*	Return array of chunks from colx
		*/
			res := Map()
			for key,val in colx {
				x1 := colx[key]
				ln := (key<colx.length) 
					? colx[key+1]-x1 
					: StrLen(txt)-x1+1
				tx := SubStr(txt,x1,ln)
				cleanspace(&tx)
				res[key] := trim(tx)
			}
			return res
		}

		getCels(idx,colx) {
		/*	Read cells based on header colx coords
		*/
			cell := Map()
			loop 3 {																	; Lookahead 3 lines
				row := A_Index
				x := readRow(txtarray[idx+row-1],colx)
				if (row>1)&&(x[1]~="^\d{1,2}:\d{2}") {
					break
				}
				for key,val in x {
					tx := x[key]
					if (row=1) {
						cell[key] := tx
					} else {
						cell[key] .= tx
					}
				}
				if (x[1]~="AM|PM") {
					cell[1] := ParseDate(cell[1]).24hms
					break
				}
			}
			return {cell:cell,height:row-1}
		}
	}

	outputCSV(obj) {
	/*	First pass: find all field names
	*/
		cols := []
		for key,val in obj {
			for r in val {
				if !ObjHasValue(cols,r) {
					cols.Push(r)
				}
			}
		}

	/*	Second pass: generate CSV
	*/
		txt := "TIME"
		for val in cols {
			txt .= ",`"" val "`""
		}
		txt .= "`n"
		for key,val in obj {
			line := key
			for r in cols {
				try res := obj[key][r]
				catch {
					res := ""
				}
				line .= ",`"" res "`""
			}
			txt .= line "`n"
		}
		FileAppend(txt,"output.csv")
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

strQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	try return (var1="") ? null : RegExReplace(txt,"###",var1)
}

ObjHasValue(aObj, aValue, rx:="") {
/*	Check if aValue is contained within aObj, return index value
	If aObj is Map(), also check key:value pairs for matching value, then key names
	If (rx), compare RegEx in both directions
 */
	if (aValue="") {																	; null aValue is error
		return false
	}
	if (props := ObjOwnPropCount(aObj)) {
		aProps := aObj.OwnProps()
		aKeys := aObj.OwnProps()
		/*	Check values of property keys
		*/
		for key,val in aProps
		{
			if (compare(val)) {
				return key
			}
		}
		/*	Check property key names
		*/
		for key,val in aKeys
		{
			if (compare(key)) {
				return key
			}
		}

	} else {
		/*	Check values in object (i.e. arrays)
		*/
		for key,val in aObj
		{
			if (compare(val)) {
				return key
			}
		}
	}
	
	return false

	compare(val) {
		if (rx) {
			if (val ~= "i)" aValue) {													; val=text, aValue=RX
				return true
			}
			if (aValue ~= "i)" val) {													; aValue=text, val=RX
				return true
			}
		} else {
			if (val = aValue) {															; otherwise just string match
				return true
			}
		}
		return false
	}
}

ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-\._/]"
	date := {yyyy:"",mmm:"",mm:"",dd:"",date:""}
	time := {hr:"",min:"",sec:"",days:"",ampm:"",time:"",24hr:""}

	x := RegExReplace(x,"[,\(\)]")

	if (x~="\d{4}.\d{2}.\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z") {
		x := RegExReplace(x,"[TZ]","|")
	}
	if (x~="\d{4}.\d{2}.\d{2}T\d{2,}") {
		x := RegExReplace(x,"T","|")
	}

	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",&d) {			; 03-Jan-2015
		date.dd := zdigit(d[1])
		date.mmm := d[2]
		date.mm := zdigit(objhasvalue(mo,d[2]))
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(\d{4})[\-\.](\d{2})[\-\.](\d{2})\b",&d) {					; 2015-01-03
		date.yyyy := d[1]
		date.mm := zdigit(d[2])
		date.mmm := mo[d[2]]
		date.dd := zdigit(d[3])
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr "|\d{1,2})" dSep "(\d{1,2})" dSep "(\d{4}|\d{2})",&d) {	; Jan-03-2015, 01-03-2015
		date.dd := zdigit(d[2])
		date.mmm := objhasvalue(mo,d[1]) 
			? d[1]
			: mo[d[1]]
		date.mm := objhasvalue(mo,d[1])
			? zdigit(objhasvalue(mo,d[1]))
			: zdigit(d[1])
		date.yyyy := (d[3]~="\d{4}")
			? d[3]
			: (d[3]>50)
				? "19" d[3]
				: "20" d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr ")\s+(\d{1,2}),?\s+(\d{4})",&d) {					; Dec 21, 2018
		date.mmm := d[1]
		date.mm := zdigit(objhasvalue(mo,d[1]))
		date.dd := zdigit(d[2])
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(19\d{2}|20\d{2})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",&d)  {	; 20150103174307 or 20150103
		date.yyyy := d[1]
		date.mm := d[2]
		date.mmm := mo[d[2]]
		date.dd := d[3]
		if (d[1]) {
			date.date := d[1] "-" d[2] "-" d[3]
		}
		
		time.hr := d[5]
		time.min := d[6]
		time.sec := d[7]
		if (d[5]) {
			time.time := d[5] ":" d[6] . strQ(d[7],":###")
		}
	}

	if RegExMatch(x,"i)(\d+):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",&t) {				; 17:42 PM
		hasDays := (t[4]) ? true : false 												; 4 nums has days
		time.days := (hasDays) ? t[1] : ""
		time.hr := trim(t[1+hasDays])
		time.min := trim(t[2+hasDays]," :")
		time.sec := trim(t[3+hasDays]," :")
		if (time.min>59) {
			time.hr := floor(time.min/60)
			time.min := zDigit(Mod(time.min,60))
		}
		if (time.hr>23) {
			time.days := floor(time.hr/24)
			time.hr := zDigit(Mod(time.hr,24))
			DHM:=true
		}
		time.ampm := trim(t[5])
		time.time := trim(t[0])

		if (time.ampm="PM")&&(time.hr<12) {
			time.24hr := time.hr + 12
		} else {
			time.24hr := time.hr
		}
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, YMDHMS:date.yyyy date.mm date.dd zDigit(time.hr) zDigit(time.min) zDigit(time.sec)
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, MMDD:date.mm "/" date.dd 
			, hrmin:zdigit(time.hr) ":" zdigit(time.min)
			, days:zdigit(time.days)
			, hr:zdigit(time.hr), min:zdigit(time.min), sec:zdigit(time.sec)
			, ampm:time.ampm, time:time.time
			, DHM:zdigit(time.days) ":" zdigit(time.hr) ":" zdigit(time.min) " (DD:HH:MM)" 
			, DT:date.mm "/" date.dd "/" date.yyyy " at " zdigit(time.hr) ":" zdigit(time.min) ":" zdigit(time.sec)
			, 24hms:zDigit(time.24hr) ":" zDigit(time.min) ":" zDigit(time.sec)}
}

zDigit(x) {
; Returns 2 digit number with leading 0
	return SubStr("00" x, -2)
}

#Include includes
#Include strx2.ahk
#Include Peep.v2.ahk