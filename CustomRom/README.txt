Custom TI-85 ROM / image converter
===================================

Files:
  CustomRomBuilder.vbp  - VB6 command-line project
  modMain.bas           - command-line entry point
  modBmpToRom.bas       - BMP decoding, monochrome conversion, ROM output
  CustomRomBuilder.exe  - compiled builder
  CustomTI85.asm        - reference source only; it is NOT assembled
  CustomTI85.rom        - generated ROM (86-byte stub + 1024-byte image)
  customimage.bmp       - example input image

What the ROM does:
  - Sets up the TI-85 LCD with contrast 18/32 (0x12) and copies the
    128 x 64 monochrome image to the screen buffer at $FC00.
  - Press ON to switch the display off.
  - Press ON again to switch it back on and redraw the image.

Build the generator:
  Open CustomRomBuilder.vbp in Visual Basic 6 and compile it, or:
  VB6.EXE /MAKE CustomRomBuilder.vbp /OUT build.log

Use the generator:
  CustomRomBuilder.exe <input.bmp> [output.rom]

Examples:
  CustomRomBuilder.exe customimage.bmp CustomTI85.rom
  CustomRomBuilder.exe "C:\Pictures\screen.bmp" "C:\Temp\test.rom"

If the output path is omitted, the builder writes CustomTI85.rom
next to the executable. There is no form; the input is taken from
Command$.

BMP support:
  - Uncompressed 1, 4, 8, 24 and 32 bpp BMP files.
  - 128 x 64 images are used directly.
  - Larger images are scaled down to fit, preserving aspect ratio.
  - Smaller images are centred at 1:1.
  - Pixels darker than luminance 160 become TI-85 "on" pixels.

The generated file is not padded to 32 KB. Ti85LoadRom zero-pads
short ROM images in memory, so only the stub and the image data
are stored.

The ASM file is kept for reference. The VB6 project writes the
machine code bytes directly; it does not invoke an assembler.
