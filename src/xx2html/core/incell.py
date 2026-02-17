"""Generate CSS rules for Excel in-cell rich-value images."""

from io import BytesIO
import logging
from zipfile import ZipFile
from PIL import Image, UnidentifiedImageError
from xlsx2html.utils.image import bytes_to_datauri

from xx2html.core.types import CellDimensions


def get_incell_css(
    vm_ids: set[str],
    vm_ids_dimension_references: dict[str, CellDimensions],
    vm_cell_vm_ids: dict[str, str],
    incell_images_refs: dict[str | None, str],
    archive: ZipFile,
) -> str:
    """Create per-vm and per-cell CSS rules for in-cell image rendering."""
    styles = []
    archive_namelist = archive.namelist()
    image_specs: dict[str, dict[str, object]] = {}

    for vm_id in sorted(vm_ids):
        rel_id = f"rId{vm_id}"

        # logging.debug(f"Transform (wb|incell): Working with vm(rId): {rel_id}")
        logging.info(f"Transform (wb|incell): Working with vm(rId): {rel_id}")
        # if rel_id in incell_images:
        # if vm_id in incell_images_refs:
        target_path = incell_images_refs.get(vm_id, None)
        if target_path is None:
            logging.warning(
                f"get_incell_css: vm(rId) -> {rel_id} not found in incell images references!"
            )
            continue
        if target_path not in archive_namelist:
            logging.warning(
                f"get_incell_css: vm(rId) -> image '{target_path}' not found in excel file!"
            )
            continue
        # logging.debug(f"Transform (wb|incell): Reading vm(rId): {rel_id} -> file at {target_path}")
        logging.info(f"get_incell_css: Reading vm(rId): {rel_id} -> file '{target_path}'")
        with archive.open(target_path) as ifile:
            image_bytes = ifile.read()

        src = bytes_to_datauri(BytesIO(image_bytes), target_path)
        width = None
        height = None
        try:
            with Image.open(BytesIO(image_bytes)) as image:
                width, height = image.size
        except (UnidentifiedImageError, OSError) as image_exc:
            logging.warning(
                "get_incell_css: Unable to read image size for vm(rId)=%s (%s): %r",
                rel_id,
                target_path,
                image_exc,
            )
        image_specs[vm_id] = {"src": src, "width": width, "height": height}

        styles.append(
            """
.vm-richvaluerel_rid%s img {
    content:url("%s");
    display:block;
}
                """
            % (vm_id, src)
        )

    for cell_class_name, cell_dimensions in vm_ids_dimension_references.items():
        vm_id = vm_cell_vm_ids.get(cell_class_name)
        image_spec = image_specs.get(vm_id, {})
        cell_width = cell_dimensions.get("width", 0)
        cell_height = cell_dimensions.get("height", 0)
        img_width = image_spec.get("width")
        img_height = image_spec.get("height")

        width_rule = "width: 100%;"
        height_rule = "height: 100%;"
        aspect_ratio_rule = ""

        if (
            isinstance(cell_width, int)
            and isinstance(cell_height, int)
            and cell_width > 0
            and cell_height > 0
            and isinstance(img_width, int)
            and isinstance(img_height, int)
            and img_width > 0
            and img_height > 0
        ):
            cell_aspect_ratio = cell_width / cell_height
            image_aspect_ratio = img_width / img_height
            if image_aspect_ratio >= cell_aspect_ratio:
                width_rule = "width: 100%;"
                height_rule = "height: auto;"
            else:
                width_rule = "width: auto;"
                height_rule = "height: 100%;"
            aspect_ratio_rule = f"aspect-ratio: {img_width} / {img_height};"

        styles.append(
            """
.%s img {
    %s
    %s
    max-width: 100%%;
    max-height: 100%%;
    object-fit: contain;
    %s
}
            """
            % (cell_class_name, width_rule, height_rule, aspect_ratio_rule)
        )

    return "\n".join(styles)
