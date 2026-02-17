from copy import deepcopy
from bs4 import BeautifulSoup


def update_links(
    html: str,
    encoded_sheet_names: dict[str, str],  # sheet names
    update_local_links: bool = True,
    update_ext_links: bool = True,
) -> str:
    soup = BeautifulSoup(html, "lxml")

    for anchor_tag in soup.find_all("a"):
        if "class" in anchor_tag.attrs and "xlsx_sheet-link" in anchor_tag["class"]:
            continue

        href = anchor_tag.get("href")
        if not isinstance(href, str) or href == "":
            continue

        is_local_anchor = href.startswith("#")

        if is_local_anchor:
            if update_local_links:
                sheet_name = href.split(".")[0].replace("#", "")
                enc_sheet_name = encoded_sheet_names.get(sheet_name)
                # sheet_name = ''.join(e for e in possible_sheet_name if e.isalnum())

                # final_href = f'#sheet-{sheet_name}' if sheet_name in usable_names else ''
                final_href = f"#{enc_sheet_name}" if enc_sheet_name is not None else ""

                sharepoint_anchor_tag = soup.new_tag(
                    "a",
                    attrs={
                        "href": f"about:srcdoc{final_href}",
                        "class": "xlsx_sheet-link sharepoint_visible",
                    },
                )

                js_anchor_tag = soup.new_tag(
                    "a",
                    attrs={
                        "href": final_href,
                        "class": "xlsx_sheet-link js_visible",
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
            if update_ext_links:
                external_anchor_tag = soup.new_tag(
                    "a",
                    attrs={
                        "href": href,
                        "class": "xlsx_sheet-link js_visible",
                        "target": "_blank",
                        "rel": "noopener noreferrer",
                    },
                )
                for child in anchor_tag.contents:
                    external_anchor_tag.append(deepcopy(child))
                anchor_tag.replace_with(external_anchor_tag)
    return str(soup)
