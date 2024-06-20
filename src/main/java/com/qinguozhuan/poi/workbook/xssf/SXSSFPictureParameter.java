package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.PictureParameter;
import com.qinguozhuan.poi.workbook.utils.ImageUtil;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import org.springframework.util.Assert;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.Serial;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Objects;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 00:18
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public final class SXSSFPictureParameter implements PictureParameter {
    /**
     * use serialVersionUID from JDK 1.0.2 for interoperability
     */
    @Serial
    private static final long serialVersionUID = -8742448824652078965L;

    private final Picture picture;

    private record Picture(byte[] bytes, int format) {
    }

    /**
     * build Object by URL
     *
     * @param url the url
     * @return SXSSFPictureParameter
     * @throws IOException
     */
    public static SXSSFPictureParameter byURL(String url) throws IOException {
        Assert.hasLength(url, "url must not be empty");

        byte[] bytes = getImageData(url);
        return byBytes(bytes);
    }

    /**
     * build Object by byte[]
     *
     * @param bytes Picture bytes data
     * @return SXSSFPictureParameter
     */
    public static SXSSFPictureParameter byBytes(byte[] bytes) {
        Objects.requireNonNull(bytes, "bytes must not be empty");
        Assert.isTrue(bytes.length > 0, "bytes must not be empty");

        int format = ImageUtil.getPicType(bytes).getOoxmlId();//{@link PictureType#getOoxmlId()}
        return new SXSSFPictureParameter(new Picture(bytes, format));
    }

    @Override
    public byte[] getBytes() {
        return picture.bytes();
    }

    @Override
    public int getFormat() {
        return picture.format();
    }

    /**
     * @param urlString picture url
     * @return Picture bytes data
     * @throws IOException
     */
    private static byte[] getImageData(String urlString) throws IOException {
        URL url = new URL(urlString);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");

        int responseCode = connection.getResponseCode();
        if (responseCode == HttpURLConnection.HTTP_OK) {
            try (InputStream inputStream = connection.getInputStream()) {
                //buffer
                byte[] buffer = new byte[1024];
                int bytesRead;
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

                // read
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    outputStream.write(buffer, 0, bytesRead);
                }

                return outputStream.toByteArray();
            }
        } else {
            throw new IOException("Server returned non-OK status: " + responseCode);
        }
    }

}
