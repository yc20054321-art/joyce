set -e
cd "$(dirname "$0")"
V=/root/.claude/skills/synced/5de0090b-eddd-4a05-af80-69d4efcd8be9_beda3221-4cc1-47b5-9233-99a60d655c1f/pptx/scripts/office/validate.py

echo "== 業務工作範疇說明 (開發／大貨／主採) =="
rm -rf unpacked
python3 -c "import zipfile; zipfile.ZipFile('src.pptx').extractall('unpacked')"
python3 gen.py
python3 renumber.py
cd unpacked && rm -f ../out.pptx && zip -qXr ../out.pptx . && cd ..
python3 $V out.pptx --original src.pptx

echo "== V6副採工作範疇簡略說明 (獨立檔) =="
python3 fucai_standalone.py
python3 $V fucai_out.pptx --original fucai_src.pptx
