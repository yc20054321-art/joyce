set -e
cd "$(dirname "$0")"
rm -rf unpacked
python3 -c "import zipfile; zipfile.ZipFile('src.pptx').extractall('unpacked')"
python3 gen.py
python3 renumber.py
cd unpacked && rm -f ../out.pptx && zip -qXr ../out.pptx . && cd ..
python3 /root/.claude/skills/synced/5de0090b-eddd-4a05-af80-69d4efcd8be9_beda3221-4cc1-47b5-9233-99a60d655c1f/pptx/scripts/office/validate.py out.pptx --original src.pptx
