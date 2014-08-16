#!/bin/bash

# Test script.

for fl in utility_test/*.txt; do 
txt2xlsx.py $fl ${fl/.txt/.xlsx} 
done

