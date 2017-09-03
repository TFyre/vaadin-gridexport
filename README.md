# vaadin-gridexport Add-on for Vaadin 8

GridExport is a Data Component add-on for Vaadin 8.

## Authors

(https://github.com/mletenay/tableexport-for-vaadin) - Inherited from
(https://github.com/jnash67/tableexport-for-vaadin) - Originally Inherited from

## Overview

* This add-on requires the Vaadin 8 library.
* This add-on requires the Apache POI library.

## Usage

```java
    protected Button getExportButton() {
        final Button button = new Button("Export XLS");
        final FileDownloader fdl = new FileDownloader(new StreamResource(() -> {
            final ExcelExport<BEAN> ee = new ExcelExport<>(new DefaultGridHolder<>(grid));
            ee.setReportTitle("Report Title");
            return ee.getInputStream();
        }, String.format("%s.xls", "Report Title")));
        fdl.extend(button);
        return button;
    }
```


## License

Add-on is distributed under Apache License 2.0. For license terms, see LICENSE.txt.