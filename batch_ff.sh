#!/bin/bash

input_full_path=/mnt/input/SIVR-040
echo "$input_full_path"
input_path=`basename $input_full_path`
echo "$input_path"

output_full_path="/mnt/output/$input_path"
echo "$output_full_path"

mkdir -p $output_full_path

cd $input_full_path

for i in *.mp4;
  do name=`echo "$i" | cut -d'.' -f1`
  echo "$i"
  echo "$name"
  echo "$output_full_path/$i"
  ffmpeg -i "$i" -c:v libx264 -r 30 -c:a copy "$output_full_path/$i"
done