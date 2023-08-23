getDaysActual(h, m) {
    if (h < 24) {
        return 1
    } else if (Mod(h, 24) = 0 && m = 0) {
        return h / 24
    } else if (h >= 24 || m != 0) {
        return h / 24 + 1
    }
}

