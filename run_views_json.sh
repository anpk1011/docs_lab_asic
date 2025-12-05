#!/usr/bin/env bash
set -euo pipefail

BASE_URL="http://192.168.1.225:8081"
AUTH="admin:Asic@2019"
LOCALE="en_us"

OUTPUT="defects.json"
> "$OUTPUT"

jq -c '.[]' projects.json | while IFS= read -r proj; do
  proj_key=$(echo "$proj" | jq -r '.proj_key')
  proj_name=$(echo "$proj" | jq -r '.proj_name')

  jq -c '.[]' views.json | while IFS= read -r view; do
    view_id=$(echo "$view" | jq -r '.view_id')
    view_name=$(echo "$view" | jq -r '.view_name')

    echo "Processing $proj_name ($proj_key) view $view_id ..."

    # Call API
    response=$(curl -s --fail --location \
      "${BASE_URL}/api/v2/views/viewContents/${view_id}?locale=${LOCALE}&projectId=${proj_key}&rowCount=-1" \
      --header 'Accept: application/json' \
      --user "${AUTH}" || true)

    #  skip if res empty
    if [[ -z "$response" ]]; then
      echo "WARNING: Empty response for $proj_name view $view_id" >&2
      continue
    fi

    # parse JSON
    jq -c \
      --arg proj "$proj_name" \
      --arg pkey "$proj_key" \
      --arg vid "$view_id" \
      --arg vname "$view_name" \
      '
      {
        project: $proj,
        project_key: $pkey,
        view_id: $vid,
        view_name: $vname,
        rows: (
          .rows | map(
            reduce .[] as $c ({}; . + {($c.key): $c.value})
          )
        )
      }
      ' <<< "$response" >> "$OUTPUT"

  done
done

echo "DONE → defects.json"
