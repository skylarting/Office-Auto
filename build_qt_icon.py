"""Generate the Windows icon used by the PySide6 mapping tool."""

from pathlib import Path

from PIL import Image, ImageDraw


def draw_icon(size: int) -> Image.Image:
    scale = size / 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    def box(values):
        return tuple(round(value * scale) for value in values)

    draw.rounded_rectangle(box((20, 20, 236, 236)), radius=round(46 * scale), fill="#1769E0")
    draw.rounded_rectangle(box((49, 48, 154, 208)), radius=round(12 * scale), fill="#FFFFFF")
    for x in (84, 119):
        draw.line(box((x, 58, x, 198)), fill="#B7D2FA", width=max(1, round(7 * scale)))
    for y in (96, 139, 182):
        draw.line(box((59, y, 144, y)), fill="#B7D2FA", width=max(1, round(7 * scale)))

    arrow_width = max(2, round(13 * scale))
    draw.line(box((134, 92, 204, 92)), fill="#FFFFFF", width=arrow_width)
    draw.polygon([box((204, 70)), box((231, 92)), box((204, 114))], fill="#FFFFFF")
    draw.line(box((218, 160, 148, 160)), fill="#FFFFFF", width=arrow_width)
    draw.polygon([box((148, 138)), box((121, 160)), box((148, 182))], fill="#FFFFFF")
    return image


def main() -> None:
    destination = Path(__file__).with_name("assets") / "excel-mapper.ico"
    destination.parent.mkdir(parents=True, exist_ok=True)
    base = draw_icon(256)
    base.save(
        destination,
        format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )


if __name__ == "__main__":
    main()
