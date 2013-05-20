import argparse

parser = argparse.ArgumentParser(description='Cleanup the Header Row of a CSV file')
parser.add_argument('-f', help='CSV file to process', dest='file')
parser.add_argument('-s', help='Remove Spaces', dest='remove_spaces')
parsr
parser.print_help()
args = parser.parse_args()
file = args.file
saveDirectory = args.output_dir

lines = open()
