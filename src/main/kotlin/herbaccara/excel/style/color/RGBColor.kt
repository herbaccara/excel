package herbaccara.excel.style.color

data class RGBColor(val red: Byte, val green: Byte, val blue: Byte) {

    constructor(red: Int, green: Int, blue: Int) : this(red.toByte(), green.toByte(), blue.toByte()) {
        if (listOf(red, green, blue).any { it !in 0..255 }) {
            throw IllegalArgumentException(String.format("Wrong RGB(%s, %s, %s)", red, green, blue))
        }
    }

    fun byteArray(): ByteArray {
        return byteArrayOf(red, green, blue)
    }
}
