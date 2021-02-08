#!/bin/bash

set -eux

# Setup the docker images
PD_DOCKER_IMG=pandoc/core
LO_DOCKER_IMG=ipunktbs/docker-libreoffice-headless

#detect platform that we're running on...
unameOut="$(uname -s)"
case "${unameOut}" in
    Linux*)     machine=Linux;;
    Darwin*)    machine=Mac;;
    CYGWIN*)    machine=Cygwin;;
    MINGW*)     machine=MinGw;;
    *)          machine="UNKNOWN:${unameOut}"
esac

# this is the core routine to process one file...
convertOne() {
	# make sure we have the docker images
	if [[ "$(docker images -q "${PD_DOCKER_IMG}" 2> /dev/null)" == "" ]]; then
		echo "Pulling Pandoc Docker image"
		docker pull "${PD_DOCKER_IMG}"
	fi
	if [[ "$(docker images -q "${LO_DOCKER_IMG}" 2> /dev/null)" == "" ]]; then
		echo "Pulling Another LibreOffice Docker image"
		docker pull "${LO_DOCKER_IMG}"
	fi

	curPath=`pwd`
	echo "curPath = ${curPath}"
	if [ "${machine}" == "MinGw" ]; then
		curPath=/`pwd`
	fi

	# Do Conversions
	echo "Converting Mardown to Word"
	docker run --rm -v "${curPath}":"${curPath}" -w "${curPath}" "${PD_DOCKER_IMG}" --defaults ./common/2docx.yml --no-highlight -o "$2".docx "$1" 

	echo "Converting Word to PDF"
	echo "$1 -> $2.pdf"
	docker run --rm -it -v "${curPath}":"${curPath}" -w "${curPath}" --name libreoffice-headless "${LO_DOCKER_IMG}" --headless --convert-to pdf:writer_pdf_Export --outdir `pwd` "$2".docx
}

# For each file specified on the command line...
for fullpath in "$@"
do
    filename="${fullpath##*/}"                      # Strip longest match of */ from start
    dir="${fullpath:0:${#fullpath} - ${#filename}}" # Substring from 0 thru pos of filename
    base="${filename%.[^.]*}"                       # Strip shortest match of . plus at least one non-dot char from end
    ext="${filename:${#base} + 1}"                  # Substring from len of base thru end
    if [[ -z "$base" && -n "$ext" ]]; then          # If we have an extension and no base, it's really the base
        base=".$ext"
        ext=""
    fi

    echo -e "$fullpath:\n\tdir  = \"$dir\"\n\tbase = \"$base\"\n\text  = \"$ext\""

	convertOne "${filename}" "${base}"
done
