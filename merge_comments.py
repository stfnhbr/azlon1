"""
Join a task's comment threads to the video annotations they're pinned to.

Both inputs are the raw GraphQL responses you copied out of the DevTools
Network tab (right-click the request -> Copy -> Copy response).

Usage:
    python merge_comments.py comments.json task.json

Writes comments_matched.csv next to the script and prints a summary.
"""

import csv
import json
import sys


def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def find_key(blob, key):
    """Walk any nested list/dict and return the first value stored under `key`."""
    if isinstance(blob, dict):
        if key in blob:
            return blob[key]
        for value in blob.values():
            found = find_key(value, key)
            if found is not None:
                return found
    elif isinstance(blob, list):
        for item in blob:
            found = find_key(item, key)
            if found is not None:
                return found
    return None


def timecode(frame, fps):
    """Frame number -> M:SS.mmm"""
    total = frame / fps
    minutes, seconds = divmod(total, 60)
    return f"{int(minutes)}:{seconds:06.3f}"


def main(comments_path, task_path):
    widget = find_key(load(task_path), "videoAnnotationWidget")["output"]
    fps = widget["fps"]

    # id -> readable name, so we can show "Male voice 2." instead of a UUID
    track_names = {t["id"]: t["name"].strip() for t in widget["tracks"]}

    annotations = {}
    for ann in widget["annotations"]:
        span = ann["spans"][0]
        values = ann["formContent"]["values"]
        annotations[ann["id"]] = {
            "track": track_names.get(ann["trackId"], ann["trackId"]),
            "start_frame": span["startFrame"],
            "end_frame": span["endFrame"],
            "start": timecode(span["startFrame"], fps),
            "end": timecode(span["endFrame"], fps),
            "caption": span["text"].strip(),
            "visibility": values.get("source_visibility", ""),
            "prominence": values.get("prominence", ""),
        }

    rows = []
    orphans = []
    for thread in find_key(load(comments_path), "searchCommentThreads")["items"]:
        for c in thread["comments"]:
            if c["deletedAt"]:          # deleted comments still come back in the JSON
                continue
            ann_id = c["location"]["videoAnnotationId"]
            ann = annotations.get(ann_id)
            if ann is None:
                orphans.append(c["comment"])
                continue
            rows.append({
                "start": ann["start"],
                "end": ann["end"],
                "track": ann["track"],
                "caption": ann["caption"],
                "visibility": ann["visibility"],
                "prominence": ann["prominence"],
                "comment": c["comment"].strip(),
                "posted_utc": c["createdAt"][:19],
                "annotation_id": ann_id,
                "_sort": ann["start_frame"],
            })

    rows.sort(key=lambda r: (r["_sort"], r["posted_utc"]))
    for r in rows:
        del r["_sort"]

    with open("comments_matched.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)

    commented = {r["annotation_id"] for r in rows}
    print(f"{len(rows)} live comments matched to {len(commented)} annotations")
    print(f"{len(annotations) - len(commented)} annotations have no comments:")
    for ann_id, ann in annotations.items():
        if ann_id not in commented:
            print(f"  {ann['start']}-{ann['end']}  {ann['track']}")
    if orphans:
        print(f"WARNING: {len(orphans)} comments pointed at unknown annotations")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
