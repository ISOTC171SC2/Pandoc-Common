#!/bin/bash

set -eux

# Setup the docker images
PD_DOCKER_IMG=pandoc/core
LO_DOCKER_IMG=ipunktbs/docker-libreoffice-headless

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

	# Do Conversions
	echo "Converting DocBook to Word"
	# docker run --rm -v `pwd`:`pwd` -w `pwd`/output "${PD_DOCKER_IMG}" -f docbook -t docx "${BASE_NAME}".xml -o "${BASE_NAME}".docx

	echo "Converting Word to PDF"
	echo "$1 -> $2.pdf"
	docker run --rm -it -v `pwd`:`pwd` -w `pwd` --name libreoffice-headless "${LO_DOCKER_IMG}" --headless --convert-to pdf:writer_pdf_Export --outdir `pwd` "$1"
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

# if you forgot the file to be processed, throw an error
# if [ -z "$1" ]
# then
#    echo "You forgot to specify which .md file to be processed";
#    exit 1
# fi
