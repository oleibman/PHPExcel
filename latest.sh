#!/bin/sh
ls -l `find Classes -type f -printf "%T@ %p\n" | sort -n | cut -d' ' -f 2- | tail -n 1`
ls -l `find Documentation -type f -printf "%T@ %p\n" | sort -n | cut -d' ' -f 2- | tail -n 1`
ls -l `find Examples -type f -printf "%T@ %p\n" | sort -n | cut -d' ' -f 2- | tail -n 1`

