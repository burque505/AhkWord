# AhkWord's documentation: MkDocs with the Material theme (mkdocs.yml, the pages in docs/).
#   python make_docs.py          writes docs/automation/api.md, the Automation API reference, from AhkWord.cs: each
#                                public member of EditorApi (what scripts call as Rte.<Name>) with its /// comment,
#                                under the headings of the class's "// ----" sections
#   python make_docs.py build    that, then builds the site into site/ (open site/index.html)
#   python make_docs.py serve    that, then serves it at http://127.0.0.1:8000, rebuilt as the pages change
# Needs mkdocs-material (pip install mkdocs-material) for build and serve; the reference alone needs only Python 3.
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(ROOT, "AhkWord.cs")
OUT = os.path.join(ROOT, "docs", "automation", "api.md")

TYPES = {"string": "text", "string?": "text", "bool": "1/0", "int": "number", "float": "number", "double": "number",
         "Color": "color", "void": ""}


def ahk_type(t):
    t = t.strip()
    return TYPES.get(t, t)


def split_params(params):
    # parameters separated by commas outside brackets and quotes
    parts, depth, quoted, current = [], 0, False, ""
    for ch in params:
        if ch == '"':
            quoted = not quoted
        if not quoted and ch in "(<[":
            depth += 1
        if not quoted and ch in ")>]":
            depth -= 1
        if ch == "," and depth == 0 and not quoted:
            parts.append(current)
            current = ""
        else:
            current += ch
    if current.strip():
        parts.append(current)
    return [p.strip() for p in parts]


def param_text(p):
    # "string path = \"\"" -> "path (text) = """
    default = ""
    if "=" in p:
        p, default = p.split("=", 1)
        default = " = " + default.strip().replace("true", "1").replace("false", "0")
    words = p.split()
    name, kind = words[-1], " ".join(words[:-1])
    kind = re.sub(r"^(params|ref|out|in)\s+", "", kind)
    return f"{name}{default}", ahk_type(kind)


def section_title(line):
    t = line.strip().lstrip("/").strip().lstrip("-").strip()
    t = re.sub(r"^Iteration \d+:\s*", "", t)
    t = re.sub(r"\s*\((?:iteration \d+|rte_\w+\.cs)\)", "", t)
    t = re.sub(r"\s*\(.*$", "", t)        # (explanations in brackets, some running on to the next line)
    t = t.split(":")[0].strip()
    return t[:1].upper() + t[1:] if t else "General"


def escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def members():
    lines = open(SOURCE, encoding="utf-8").read().splitlines()
    start = next(i for i, l in enumerate(lines) if l.startswith("public class EditorApi"))
    sections, current, comment = [], ("General", []), []
    sections.append(current)

    def body(at):
        # a member's lines, from its first to where its braces close (or the line alone, for "=>" and "{ get; set; }")
        depth, out = 0, []
        for l in lines[at:]:
            out.append(l)
            depth += l.count("{") - l.count("}")
            if depth <= 0 and ("{" in "".join(out) or l.rstrip().endswith(";")):
                break
        return "\n".join(out)

    for index, line in enumerate(lines[start + 1:], start + 1):
        if re.match(r"^(public |internal |sealed |static )*(class|enum|record|struct) ", line):
            break    # the next type: the class has ended
        if line.startswith("\t// ----"):
            title = section_title(line)
            if title == "Helpers":
                current = ("Formatting helpers", [])
            else:
                current = (title, [])
            sections.append(current)
            comment = []
            continue
        if line.startswith("\t///"):
            comment.append(line.strip()[3:].strip())
            continue
        m = re.match(r"^\tpublic (?!event |const |static |EditorApi\()(.+?)\s*(\{|=>|$)", line)
        if m:
            sig = m.group(1).strip()
            if "(" in sig and (sig.index("(") < sig.find("{") if "{" in sig else True):
                head, params = sig.split("(", 1)
                params = params.rsplit(")", 1)[0]
                *kind, name = head.split()
                current[1].append(("method", name, ahk_type(" ".join(kind)), [param_text(p) for p in split_params(params)], " ".join(comment)))
            else:
                *kind, name = sig.split()
                whole = body(index)
                setter = re.search(r"(?<![\w.])set\s*(;|\{|=>)", whole) is not None and not re.search(r"\b(private|internal)\s+set\b", whole)
                current[1].append(("property", name, ahk_type(" ".join(kind)), setter, " ".join(comment)))
        if line.strip() and not line.startswith("\t///"):
            comment = [] if not line.startswith("\t\t") and not line.startswith("\t[") else comment
    return [s for s in sections if s[1]]


def write_reference():
    out = ["# Automation API reference", "",
           "Everything a script can call on AhkWord, as `Rte.<Name>` after `#Include \"RteClient2.ahk\"` (see "
           "[Automation from AutoHotkey](index.md)). Methods are called `Rte.Name(arguments)`; properties are read as "
           "`Rte.Name` and set as `Rte.Name := value`.", "",
           "Types: *text*; *number*; *1/0* for yes or no; enumerations by name (`Heading1`, `LowerAlpha`); colours as "
           "`#RRGGBB` or a name. An argument with `= value` can be left out.", "",
           "!!! note \"Made from the program\"",
           "    This page is generated from AhkWord's source by `make_docs.py`, so it always matches the program. "
           "Tools > Automation API in AhkWord lists the same members, and the editor control's own as well.", ""]
    for title, items in members():
        out += [f"## {escape(title)}", ""]
        for item in sorted(items, key=lambda i: i[1].lower()):
            if item[0] == "method":
                _, name, returns, params, doc = item
                args = ", ".join(p for p, _ in params)
                out.append(f"### {name}")
                out.append("")
                out.append(f"```autohotkey\nRte.{name}({args})\n```")
                facts = [f"*{p.split(' = ')[0]}*: {t}" for p, t in params if t]
                if returns:
                    facts.append(f"returns: {returns}")
                if facts:
                    out.append("")
                    out.append(" · ".join(facts))
            else:
                _, name, kind, setter, doc = item
                out.append(f"### {name}")
                out.append("")
                out.append(f"```autohotkey\nvalue := Rte.{name}" + (f"\nRte.{name} := value" if setter else "") + "\n```")
                out.append("")
                out.append(f"{kind}" + ("" if setter else " · read-only"))
            if doc:
                out += ["", escape(doc)]
            out.append("")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out))
    return sum(len(s[1]) for s in members())


if __name__ == "__main__":
    count = write_reference()
    print(f"docs/automation/api.md: {count} members")
    if len(sys.argv) > 1 and sys.argv[1] in ("build", "serve"):
        sys.exit(subprocess.call([sys.executable, "-m", "mkdocs", sys.argv[1]], cwd=ROOT))
