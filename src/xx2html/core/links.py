"""Rewrite worksheet and external links in generated HTML."""

from copy import deepcopy
from urllib.parse import urlparse

from bs4 import BeautifulSoup


def _normalize_space_tokens(value: object) -> list[str]:
    if isinstance(value, str):
        return [token for token in value.split() if token]
    if isinstance(value, (list, tuple, set)):
        tokens: list[str] = []
        for item in value:
            if isinstance(item, str):
                tokens.extend([token for token in item.split() if token])
        return tokens
    return []


def _merge_tokens(value: object, *required_tokens: str) -> str:
    merged_tokens: list[str] = []
    for token in _normalize_space_tokens(value) + list(required_tokens):
        if token and token not in merged_tokens:
            merged_tokens.append(token)
    return " ".join(merged_tokens)


def _is_rewritable_external_href(href: str) -> bool:
    if href.startswith("//"):
        return True
    parsed_href = urlparse(href)
    return parsed_href.scheme in {"http", "https"}


def _stringify_attr_value(value: object) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, (list, tuple, set)):
        parts: list[str] = []
        for item in value:
            if isinstance(item, str):
                parts.append(item)
            else:
                parts.append(str(item))
        return " ".join(part for part in parts if part)
    return str(value)


def _collect_base_attrs(anchor_tag, excluded_keys: set[str]) -> dict[str, str]:
    base_attrs: dict[str, str] = {}
    for key, value in anchor_tag.attrs.items():
        if key in excluded_keys:
            continue
        base_attrs[str(key)] = _stringify_attr_value(deepcopy(value))
    return base_attrs


def update_links(
    html: str,
    encoded_sheet_names: dict[str, str],  # sheet names
    update_local_links: bool = True,
    update_ext_links: bool = True,
) -> str:
    """Rewrite anchor tags for worksheet-local and external navigation."""
    soup = BeautifulSoup(html, "lxml")
    update_links_in_soup(
        soup,
        encoded_sheet_names,
        update_local_links=update_local_links,
        update_ext_links=update_ext_links,
    )
    return str(soup)


def update_links_in_soup(
    soup: BeautifulSoup,
    encoded_sheet_names: dict[str, str],
    update_local_links: bool = True,
    update_ext_links: bool = True,
) -> None:
    """Rewrite links in-place on an existing parsed HTML soup."""

    def resolve_sheet_name(local_reference: str) -> str:
        """Resolve local reference text to a worksheet name key."""
        if local_reference in encoded_sheet_names:
            return local_reference

        if "!" in local_reference:
            sheet_name = local_reference.split("!", 1)[0]
            if sheet_name.startswith("'") and sheet_name.endswith("'"):
                return sheet_name[1:-1].replace("''", "'")
            return sheet_name

        if "." in local_reference:
            candidate = local_reference.rsplit(".", 1)[0]
            if candidate in encoded_sheet_names:
                return candidate

        return local_reference

    for anchor_tag in soup.find_all("a"):
        current_classes = _normalize_space_tokens(anchor_tag.get("class"))
        if "xlsx_sheet-link" in current_classes:
            continue

        href = anchor_tag.get("href")
        if not isinstance(href, str) or href == "":
            continue

        is_local_anchor = href.startswith("#")

        if is_local_anchor:
            if update_local_links:
                local_reference = href[1:]
                sheet_name = resolve_sheet_name(local_reference)
                enc_sheet_name = encoded_sheet_names.get(sheet_name)
                if enc_sheet_name is None:
                    continue
                # sheet_name = ''.join(e for e in possible_sheet_name if e.isalnum())

                # final_href = f'#sheet-{sheet_name}' if sheet_name in usable_names else ''
                final_href = f"#{enc_sheet_name}"
                base_attrs = _collect_base_attrs(
                    anchor_tag, {"href", "class", "target", "rel"}
                )

                sharepoint_anchor_tag = soup.new_tag(
                    "a",
                    attrs=base_attrs
                    | {
                        "href": f"about:srcdoc{final_href}",
                        "class": _merge_tokens(
                            anchor_tag.get("class"),
                            "xlsx_sheet-link",
                            "sharepoint_visible",
                        ),
                    },
                )

                js_anchor_tag = soup.new_tag(
                    "a",
                    attrs=base_attrs
                    | {
                        "href": final_href,
                        "class": _merge_tokens(
                            anchor_tag.get("class"), "xlsx_sheet-link", "js_visible"
                        ),
                        # , 'target': '_blank'
                    },
                )

                for child in anchor_tag.contents:
                    sharepoint_anchor_tag.append(deepcopy(child))
                    js_anchor_tag.append(deepcopy(child))

                # print(f"sharepoint: {sharepoint_anchor_tag}")
                # print(f"js: {js_anchor_tag}")

                anchor_tag.replace_with(sharepoint_anchor_tag, js_anchor_tag)
        else:
            if update_ext_links and _is_rewritable_external_href(href):
                base_attrs = _collect_base_attrs(
                    anchor_tag, {"href", "class", "target", "rel"}
                )
                target = anchor_tag.get("target")
                resolved_target = (
                    target if isinstance(target, str) and target else "_blank"
                )
                external_anchor_tag = soup.new_tag(
                    "a",
                    attrs=base_attrs
                    | {
                        "href": href,
                        "class": _merge_tokens(
                            anchor_tag.get("class"), "xlsx_sheet-link", "js_visible"
                        ),
                        "target": resolved_target,
                        "rel": _merge_tokens(
                            anchor_tag.get("rel"), "noopener", "noreferrer"
                        ),
                    },
                )
                for child in anchor_tag.contents:
                    external_anchor_tag.append(deepcopy(child))
                anchor_tag.replace_with(external_anchor_tag)
