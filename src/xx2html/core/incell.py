import logging
from zipfile import ZipFile
from xlsx2html.utils.image import bytes_to_datauri


def get_incell_css(
    vm_ids: set[str],
    vm_ids_dimension_references: dict[str, str],
    incell_images_refs: dict[str | None, str],
    archive: ZipFile,
) -> str:
    styles = []
    archive_namelist = archive.namelist()

    for vm_id in vm_ids:
        rId = f"rId{vm_id}"
        
        # logging.debug(f"Transform (wb|incell): Working with vm(rId): {rId}")
        logging.info(f"Transform (wb|incell): Working with vm(rId): {rId}")
        # if rId in incell_images:
        # if vm_id in incell_images_refs:
        target_path = incell_images_refs.get(vm_id, None)
        if target_path is None:
            logging.warning(f"get_incell_css: vm(rId) -> {rId} not found in incell images references!")
            continue
        if target_path not in archive_namelist:
            logging.warning(f"get_incell_css: vm(rId) -> image '{target_path}' not found in excel file!")
            continue
        # logging.debug(f"Transform (wb|incell): Reading vm(rId): {rId} -> file at {targetPath}")
        logging.info(f"get_incell_css: Reading vm(rId): {rId} -> file '{target_path}'")
        ifile = archive.open(target_path)
        src = bytes_to_datauri(ifile, target_path)
        styles.append(
                """
.vm-richvaluerel_rid%s img {
    content:url("%s");
    display:block;
}
                """
                % (vm_id, src)
            )

    for cell_class_name in vm_ids_dimension_references:
        styles.append(
            """
.%s img {
    %s : 100%%;
}
            """
            % (cell_class_name, vm_ids_dimension_references[cell_class_name])
        )

    return "\n".join(styles)

