import colorsys


def generate_random_gradient(base_color, lightness_variation=0.1, num_colors=6):
    """
    Generate a list of HEX color codes forming a gradient based on a base HSL color.
    :param base_color: tuple (h, s, l) where h is in [0,360] and s, l in [0,1].
    """

    def hsl_to_hex(h, s, l):
        r, g, b = colorsys.hls_to_rgb(h / 360, l, s)
        return f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"

    colors = []
    for i in range(num_colors):
        new_lightness = max(
            0, min(1, base_color[2] + (i - num_colors // 2) * lightness_variation)
        )
        colors.append(hsl_to_hex(base_color[0], base_color[1], new_lightness))
    return colors
