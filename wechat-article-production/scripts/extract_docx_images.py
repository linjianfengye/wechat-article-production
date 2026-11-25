"""Export embedded DOCX images in main-document reading order."""
import argparse
from collections import Counter
import json
from pathlib import Path, PurePosixPath
import posixpath
import re
from tempfile import TemporaryDirectory
from urllib.parse import unquote
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile


def local_name(name):
    return name.rsplit("}", 1)[-1]


def relationship_attribute(node, *names):
    for name in names:
        for attribute, value in node.attrib.items():
            if local_name(attribute) == name and (
                "relationships}" in attribute or "office}relid" in attribute
            ):
                return value
    return None


def resolve_part(owner, target):
    target = unquote(target)
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))


def load_relationships(package, part):
    if not part:
        path = "_rels/.rels"
    else:
        directory, filename = posixpath.split(part)
        path = posixpath.join(directory, "_rels", filename + ".rels")
    try:
        root = ET.fromstring(package.read(path))
    except KeyError:
        return {}
    return {node.attrib["Id"]: node.attrib for node in root if "Id" in node.attrib}


def ordered_children(node):
    children = list(node)
    if local_name(node.tag) != "AlternateContent":
        return children
    choices = [item for item in children if local_name(item.tag) == "Choice"]
    fallback = [item for item in children if local_name(item.tag) == "Fallback"]
    # Branches often contain DrawingML and VML versions of the same image.
    for branch in choices + fallback:
        if any(local_name(item.tag) in {"blip", "imagedata"} for item in branch.iter()):
            return [branch]
    return (choices or fallback)[:1]


def walk_paragraphs(node, path="body"):
    name = local_name(node.tag)
    if name in {"del", "moveFrom"}:
        return
    if name == "p":
        yield path, node
        return
    indices = Counter()
    for child in ordered_children(node):
        tag = local_name(child.tag)
        indices[tag] += 1
        yield from walk_paragraphs(child, f"{path}/{tag}[{indices[tag]}]")


def image_relationship_ids(node, inside_picture=False):
    name = local_name(node.tag)
    if name in {"object", "del", "moveFrom"}:
        return
    inside_picture = inside_picture or name in {"drawing", "pict"}
    if inside_picture and name in {"blip", "imagedata"}:
        reference = relationship_attribute(node, "embed", "id", "relid", "link")
        if reference:
            yield reference
            return
    for child in ordered_children(node):
        yield from image_relationship_ids(child, inside_picture)


def visible_text(node):
    if local_name(node.tag) in {"del", "moveFrom", "drawing", "pict", "object"}:
        return ""
    if local_name(node.tag) == "t":
        return node.text or ""
    return "".join(visible_text(child) for child in ordered_children(node))


def extract(package, destination, source_name):
    root_rels = load_relationships(package, "")
    main_rel = next(
        (rel for rel in root_rels.values() if rel.get("Type", "").endswith("/officeDocument")),
        {},
    )
    main_part = resolve_part("", main_rel.get("Target", "word/document.xml"))
    document = ET.fromstring(package.read(main_part))
    body = next((node for node in document if local_name(node.tag) == "body"), None)
    if body is None:
        raise ValueError("The main document has no body.")
    relationships = load_relationships(package, main_part)
    records = []
    for location, paragraph in walk_paragraphs(body):
        for reference in image_relationship_ids(paragraph):
            index = len(records) + 1
            item = {"order": index, "location": location, "paragraph_text": visible_text(paragraph),
                    "relationship_id": reference}
            relation = relationships.get(reference, {})
            if not relation.get("Type", "").endswith("/image") or not relation.get("Target"):
                item["status"] = "missing_image_relationship"
            elif relation.get("TargetMode", "").lower() == "external":
                item["status"] = "external_image_skipped"
            else:
                source_part = resolve_part(main_part, relation["Target"])
                item["source_part"] = source_part
                try:
                    image_data = package.read(source_part)
                except KeyError:
                    item["status"] = "missing_image_part"
                else:
                    suffix = PurePosixPath(source_part).suffix.lower()
                    if not re.fullmatch(r"\.[a-z0-9]{1,10}", suffix):
                        suffix = ".bin"
                    filename = f"{index:03d}{suffix}"
                    (destination / filename).write_bytes(image_data)
                    item.update(status="exported", filename=filename)
            records.append(item)
    exported = sum(item["status"] == "exported" for item in records)
    inventory = {"source": source_name, "image_occurrences": len(records),
                 "exported": exported, "images": records}
    (destination / "image-inventory.json").write_text(
        json.dumps(inventory, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return inventory


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("output", type=Path, help="a new directory")
    args = parser.parse_args()
    try:
        if args.output.exists():
            raise FileExistsError(f"Output directory already exists: {args.output}")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory(dir=args.output.parent) as temp, ZipFile(args.docx) as package:
            staging = Path(temp) / "images"
            staging.mkdir()
            inventory = extract(package, staging, args.docx.name)
            staging.rename(args.output)
    except (OSError, BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
        parser.exit(1, f"Image export failed: {exc}\n")
    print(f"Exported {inventory['exported']} of {inventory['image_occurrences']} images.")
    print(args.output / "image-inventory.json")


if __name__ == "__main__":
    main()
