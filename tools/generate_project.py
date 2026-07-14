from __future__ import annotations

import json
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / "project.yml"


def load_project() -> dict:
    with CONFIG.open("r", encoding="utf-8") as file:
        data = yaml.safe_load(file)
    if not isinstance(data, dict) or "project" not in data:
        raise ValueError("project.yml must contain a top-level 'project' section.")
    project = data["project"]
    required = [
        "name", "description", "version", "organization", "author", "license",
        "github", "documentation", "doi", "zenodo", "releases", "issues"
    ]
    missing = [key for key in required if not project.get(key)]
    if missing:
        raise ValueError("Missing required project.yml fields: " + ", ".join(missing))
    return project


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content.strip() + "\n", encoding="utf-8")


def bibtex_key(name: str) -> str:
    return name.replace("-", "_").replace(" ", "_")


def generate_readme(p: dict) -> None:
    content = f"""# {p["name"]}

{p["description"]}

## Project Resources

| Resource | Link |
|----------|------|
| Documentation | [GitHub Pages]({p["documentation"]}) |
| Repository | [GitHub]({p["github"]}) |
| DOI | [Zenodo](https://doi.org/{p["doi"]}) |
| Latest Release | [GitHub Releases]({p["releases"]}) |
| Issue Tracker | [GitHub Issues]({p["issues"]}) |

## Version

`{p["version"]}`

## License

{p["license"]}

## SportsLabResearch

https://github.com/{p["organization"]}
"""
    write_text(ROOT / "README.md", content)


def generate_home(p: dict) -> None:
    content = f"""# {p["name"]}

{p["description"]}

## Resources

| Resource | Link |
|----------|------|
| Repository | [{p["name"]}]({p["github"]}) |
| Documentation | [GitHub Pages]({p["documentation"]}) |
| DOI | [{p["doi"]}](https://doi.org/{p["doi"]}) |
| Latest Release | [GitHub Releases]({p["releases"]}) |
| Issue Tracker | [GitHub Issues]({p["issues"]}) |

## Citation

Please cite the official Zenodo DOI when using this software:

https://doi.org/{p["doi"]}
"""
    write_text(ROOT / "docs" / "index.md", content)


def generate_citation(p: dict) -> None:
    key = bibtex_key(p["name"])
    content = f"""# Citation

Cite {p["name"]} whenever the software contributes to research or scientific publications.

## Official DOI

https://doi.org/{p["doi"]}

## Recommended Citation

**{p["organization"]}.** *{p["name"]}*. Version {p["version"]}. Zenodo.

https://doi.org/{p["doi"]}

## BibTeX

```bibtex
@software{{{key},
  author    = {{{p["organization"]}}},
  title     = {{{p["name"]}}},
  year      = {{2026}},
  version   = {{{p["version"]}}},
  publisher = {{Zenodo}},
  doi       = {{{p["doi"]}}},
  url       = {{https://doi.org/{p["doi"]}}}
}}
```
"""
    write_text(ROOT / "docs" / "citation.md", content)


def generate_citation_cff(p: dict) -> None:
    content = f"""cff-version: 1.2.0
message: "Please cite this software using the metadata below."
title: "{p["name"]}"
version: "{p["version"]}"
doi: "{p["doi"]}"
url: "{p["github"]}"
repository-code: "{p["github"]}"
license: "{p["license"]}"
authors:
  - name: "{p["author"]}"
"""
    write_text(ROOT / "CITATION.cff", content)


def generate_zenodo(p: dict) -> None:
    metadata = {
        "title": p["name"],
        "description": p["description"],
        "creators": [{"name": p["author"]}],
        "license": p["license"],
        "upload_type": "software",
        "version": p["version"],
        "related_identifiers": [
            {"identifier": p["github"], "relation": "isSupplementTo", "scheme": "url"},
            {"identifier": p["documentation"], "relation": "isDocumentedBy", "scheme": "url"},
        ],
    }
    (ROOT / ".zenodo.json").write_text(
        json.dumps(metadata, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def generate_project_json(p: dict) -> None:
    (ROOT / "project.json").write_text(
        json.dumps(p, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def main() -> None:
    project = load_project()
    generate_readme(project)
    generate_home(project)
    generate_citation(project)
    generate_citation_cff(project)
    generate_zenodo(project)
    generate_project_json(project)
    print("Generated:")
    print(" - README.md")
    print(" - docs/index.md")
    print(" - docs/citation.md")
    print(" - CITATION.cff")
    print(" - .zenodo.json")
    print(" - project.json")


if __name__ == "__main__":
    main()
