#!/bin/bash

inputDir="./"
Output="result.xlsx"
PrevModule="dtsre"
Match=""
Append=0
TokenThrd=75

print_help(){
	echo "Usage: $0 [-b] [-d] [-o] [-m] [-a] [-t] [-h]"
	echo " -b   Specify a keyword before module name, eg dtsre"
	echo " -d	Specify Input Directory, multiple dirs are seperate by ,"
	echo " -o	Output .xlsx file path"
	echo " -m	Match pattern for filter, \"A|B\"--means source duplicate from A and B"
	echo " -a	Append result to target worksheet"
	echo " -t   Minimal token threshold"
	echo " -h	print this help message"
 }

while getopts ":d:o:m:h:t:b:a" opt; do
	case $opt in
		b)
			prevModule=$OPTARG
			;;
		d)
			inputDir=$OPTARG
			;;
		o)
			Output=$OPTARG
			;;		
		m)
			Match=$OPTARG
			;;
		a)
			Append=1
			;;
		t)
			TokenThrd=$OPTARG
			;;	
		h)
			print_help
			exit 1
			;;

		*)
			echo "Error: Unknown option: $OPTARG"
			print_help
			exit 1
			;;
	esac
done

if [[ -f "tmp.txt" ]]; then
	rm "tmp.txt"
fi

cpdOptions=""

if [ ! -z "$prevModule" ]; then
	cpdOptions+="-b ${prevModule} "
fi

if [ ! -z "$Output" ]; then
	cpdOptions+="-o ${Output} "
fi

if [ ! -z "$Match" ]; then
	cpdOptions+="-m ${Match} "
fi

if [[ "$Append" -eq 1 ]]; then
	cpdOptions+="-a "
fi


#pmd.bat cpd --minimum-tokens=75 --ignore-literals --ignore-identifiers --ignore-literal-sequences --ignore-sequences -l cpp -e GBK -d ${inputDir} > tmp.txt

pmd.bat cpd --minimum-tokens=${TokenThrd} --ignore-literals --ignore-identifiers --ignore-literal-sequences --no-fail-on-error -l cpp -d ${inputDir} > tmp.txt

echo start runing CPDFilter.exe...

echo args: -d ${inputDir} ${cpdOptions}
CPDFilter.exe tmp.txt ${cpdOptions}

echo file written. ${Output}