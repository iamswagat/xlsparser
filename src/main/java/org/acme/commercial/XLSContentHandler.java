package org.acme.commercial;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.Objects;
import java.util.Optional;

@Slf4j
public class XLSContentHandler implements XSSFBSheetHandler.SheetContentsHandler {
    /**
     * A cell, with the given formatted value (may be null),
     * a url (may be null), a toolTip (may be null)
     * and possibly a comment (may be null), was encountered
     *
     * @param cellReference
     * @param formattedValue
     * @param url
     * @param toolTip
     * @param comment
     */
    @Override
    public void hyperlinkCell(String cellReference, String formattedValue, String url, String toolTip, XSSFComment comment) {
        log.info(">>>>>hyperlinkCell::start<<<<<<");
        log.info("cellReference: ".concat(cellReference));
        log.info("formattedValue: ".concat(formattedValue));
        log.info("url: ".concat(url));
        log.info("toolTip: ".concat(toolTip));
        log.info("any comment".concat(BooleanUtils.toStringTrueFalse(Objects.nonNull(comment))));
        log.info(">>>>>hyperlinkCell::start<<<<<<");
    }

    /**
     * A row with the (zero based) row number has started
     *
     * @param rowNum
     */
    @Override
    public void startRow(int rowNum) {
        log.info("startRow: " + rowNum);
    }

    /**
     * A row with the (zero based) row number has ended
     *
     * @param rowNum
     */
    @Override
    public void endRow(int rowNum) {
        log.info("endRow: " + rowNum);
    }

    /**
     * A cell, with the given formatted value (may be null),
     * and possibly a comment (may be null), was encountered
     *
     * @param cellReference
     * @param formattedValue
     * @param comment
     */
    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        log.info("cellReference:" + cellReference + ", formattedValue: " + formattedValue);
        Optional.ofNullable(comment).ifPresent(c -> {
            log.info(c.getAddress().formatAsString());
        });
    }

    /**
     * A header or footer has been encountered
     *
     * @param text
     * @param isHeader
     * @param tagName
     */
    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {

    }
}
