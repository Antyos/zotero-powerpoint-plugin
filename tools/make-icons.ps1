$sizes = (16, 32, 64, 80, 128);

foreach ($size in $sizes) {
    & inkscape --export-filename="assets/icon-${size}.png" `
        --export-area-page `
        --export-width=${size} `
        --export-height=${size} `
        "assets/logo.svg"
}
