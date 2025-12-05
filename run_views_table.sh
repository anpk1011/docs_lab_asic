#!/usr/bin/env bash
set -euo pipefail

BASE_URL="http://192.168.1.225:8081"
AUTH="admin:Asic@2019"
LOCALE="en_us"
OUTPUT_FILE="report.txt"

# Xóa file cũ nếu tồn tại
> "$OUTPUT_FILE"

function print_and_save() {
    tee -a "$OUTPUT_FILE"
}

jq -c '.[]' projects.json | while read -r proj; do
  proj_key=$(echo "$proj" | jq -r '.proj_key')
  proj_name=$(echo "$proj" | jq -r '.proj_name')

  echo -e "\n===============================" | print_and_save
  echo "PROJECT: $proj_name ($proj_key)" | print_and_save
  echo "===============================" | print_and_save

  jq -c '.[]' views.json | while read -r view; do
    view_id=$(echo "$view" | jq -r '.view_id')
    view_name=$(echo "$view" | jq -r '.view_name')

    echo -e "\n--- VIEW: $view_name ($view_id) ---" | print_and_save

    curl -s --location \
      --request GET "${BASE_URL}/api/v2/views/viewContents/${view_id}?locale=${LOCALE}&projectId=${proj_key}&rowCount=-1" \
      --header 'Accept: application/json' \
      --user "${AUTH}" \
      | jq -r '
          (.columns | map(.name) | join("\t")),
          (.rows[] | map(.value) | join("\t"))
        ' \
      | column -t \
      | print_and_save

  done
done

echo -e "\nDONE! Kết quả được lưu trong: $OUTPUT_FILE"
