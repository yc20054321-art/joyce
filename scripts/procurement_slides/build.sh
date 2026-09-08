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

echo "== V6副採工作範疇簡略說明 v2 (獨立檔) =="
python3 fucai_v2.py
python3 $V fucai_v2_out.pptx --original q1.pptx
python3 qa.py

echo "== V6主採業務範疇說明 v2 (獨立檔) =="
python3 zhucai_v2.py
python3 $V zhucai_v2_out.pptx --original g1.pptx

echo "== 主副料採購 AI 導入評估 (整合版) =="
python3 integrate.py
python3 $V integrated_out.pptx
