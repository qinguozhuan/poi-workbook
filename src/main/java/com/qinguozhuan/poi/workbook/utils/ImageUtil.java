package com.qinguozhuan.poi.workbook.utils;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import org.apache.poi.common.usermodel.PictureType;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 21:30
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public final class ImageUtil {

    /**
     * 常用文件的文件头如下：(以前六位为准)
     * JPEG (jpg)，文件头：FFD8FF
     * PNG (png)，文件头：89504E47
     * GIF (gif)，文件头：47494638
     * TIFF (tif)，文件头：49492A00
     * Windows Bitmap (bmp)，文件头：424D
     * CAD (dwg)，文件头：41433130
     * Adobe Photoshop (psd)，文件头：38425053
     * Rich Text Format (rtf)，文件头：7B5C727466
     * XML (xml)，文件头：3C3F786D6C
     * HTML (html)，文件头：68746D6C3E
     * Email [thorough only] (eml)，文件头：44656C69766572792D646174653A
     * Outlook Express (dbx)，文件头：CFAD12FEC5FD746F
     * Outlook (pst)，文件头：2142444E
     * MS Word/Excel (xls.or.doc)，文件头：D0CF11E0
     * MS Access (mdb)，文件头：5374616E64617264204A
     * WordPerfect (wpd)，文件头：FF575043
     * Postscript (eps.or.ps)，文件头：252150532D41646F6265
     * Adobe Acrobat (pdf)，文件头：255044462D312E
     * Quicken (qdf)，文件头：AC9EBD8F
     * Windows Password (pwl)，文件头：E3828596
     * ZIP Archive (zip)，文件头：504B0304
     * RAR Archive (rar)，文件头：52617221
     * Wave (wav)，文件头：57415645
     * AVI (avi)，文件头：41564920
     * Real Audio (ram)，文件头：2E7261FD
     * Real Media (rm)，文件头：2E524D46
     * MPEG (mpg)，文件头：000001BA
     * MPEG (mpg)，文件头：000001B3
     * Quicktime (mov)，文件头：6D6F6F76
     * Windows Media (asf)，文件头：3026B2758E66CF11
     * MIDI (mid)，文件头：4D546864
     *
     * @param bytes 图片文件
     * @return jpg/png/gif/bmp等枚举类型
     */
    public static PictureType getPicType(byte[] bytes) {
        String hexString = bytes2Hex(bytes).toUpperCase();
        if (hexString.startsWith("FFD8FF")) {
            return PictureType.JPEG;
        } else if (hexString.startsWith("89504E47")) {
            return PictureType.PNG;
        } else if (hexString.startsWith("47494638")) {
            return PictureType.GIF;
        } else if (hexString.startsWith("424D")) {
            return PictureType.BMP;
        } else if (hexString.startsWith("49492A00")) {
            return PictureType.TIFF;
        } else {
            return PictureType.UNKNOWN;
        }
    }

    /**
     * byte数组转换成16进制字符串， 只需要前4位
     *
     * @param bytes 图片文件
     * @return 16进制字符串
     */
    private static String bytes2Hex(byte[] bytes) {
        StringBuilder sb = new StringBuilder();
        if (bytes == null || bytes.length == 0) {
            return sb.toString();
        }
        for (int i = 0; i < Math.min(bytes.length, 4); i++) {
            int v = bytes[i] & 0xFF;
            String hv = Integer.toHexString(v);
            if (hv.length() < 2) {
                sb.append(0);
            }
            sb.append(hv);
        }
        return sb.toString();
    }

    /**
     * @param bytes image bytes
     * @return {@link BufferedImage}
     * @throws IOException {@link  IOException}
     */
    public static BufferedImage toBufferedImage(byte[] bytes) throws IOException {
        try (InputStream inputStream = new ByteArrayInputStream(bytes)) {
            return ImageIO.read(inputStream);
        }
    }

}
