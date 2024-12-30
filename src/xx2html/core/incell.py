import logging
from zipfile import ZipFile
from xlsx2html.utils.image import bytes_to_datauri


def get_incell_css(
    vm_ids: set[str],
    vm_ids_dimension_references: dict[str, str],
    incell_images: dict[str | None, str],
    archive: ZipFile,
) -> str:
    styles = []

    for vm_id in vm_ids:
        rId = f"rId{vm_id}"
        
        logging.debug(f"Transform (wb|incell): Working with vm(rId): {rId}")
        if rId in incell_images:
            targetPath = incell_images[rId]
            logging.debug(f"Transform (wb|incell): Reading vm(rId): {rId} -> file at {targetPath}")
            ifile = archive.open(targetPath)
            src = bytes_to_datauri(ifile, targetPath)
            styles.append(
                """
.vm-richvaluerel_rid%d img {
    content:url("%s");
    display:block;
}
                """
                % (vm_id, src)
            )
        else:
            logging.error(f"Transform (wb|incell): vm(rId): {rId} not found!")

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

