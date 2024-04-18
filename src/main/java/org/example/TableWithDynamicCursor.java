package org.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XDependentTextField;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextField;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;

import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XLineCursor;
import com.sun.star.view.XSelectionSupplier;
import javax.swing.JOptionPane;
import ooo.connector.BootstrapSocketConnector;

import static com.sun.star.uno.UnoRuntime.queryInterface;

public class TableWithDynamicCursor {

    static final String DOCUMENT_PATH = "document_with_table.odt";
    static final String LIBREOFFICE_PATH = "/opt/libreoffice7.3/program/";

    public static void main(String[] args) throws Exception {
        // When Cinnamon runs with args -Djdk.gtk.version=2

        // load file from project
        final String path = TableWithDynamicCursor.class.getResource(DOCUMENT_PATH).toString();

        // setup
        XComponentContext componentContext = BootstrapSocketConnector.bootstrap(LIBREOFFICE_PATH);

        XMultiComponentFactory xMCF = componentContext.getServiceManager();
        Object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", componentContext);
        XDesktop xDesktop = queryInterface(XDesktop.class, oDesktop);
        XComponentLoader xCLoader = queryInterface(XComponentLoader.class, oDesktop);

        // load example odt
        XComponent xComponent = xCLoader.loadComponentFromURL(path, "_default", 0, new PropertyValue[]{ });

        final XTextDocument xCurrentTextDocument = queryInterface(XTextDocument.class, xComponent);

        // emulates a user interaction
        int selectedOption = JOptionPane.showConfirmDialog(null, "Manually, click inside any cell and continue process");
        if ( selectedOption != 0 ) {
            xCurrentTextDocument.dispose();
            System.exit(0);
        }


        //---------------------------- wrapTableWithDynamicBlock ------------------------------
        // getCurrentCursor
        XSelectionSupplier xSelSupplier = queryInterface(XSelectionSupplier.class, xCurrentTextDocument.getCurrentController());
        Object oSelection = xSelSupplier.getSelection();
        XServiceInfo xServInfo = queryInterface(XServiceInfo.class, oSelection);
        XIndexAccess xIndexAccess = queryInterface(XIndexAccess.class, oSelection);
        XTextRange currentCursor = queryInterface(XTextRange.class, xIndexAccess.getByIndex(0));

        // getCurrentTable from current cursor
        final XPropertySet xPropertySet = queryInterface(XPropertySet.class, currentCursor);
        Object textTable = xPropertySet.getPropertyValue("TextTable");
        XTextTable xTextTable = queryInterface(XTextTable.class, textTable);

        // getRangeFromTable
        final XTextRange xStartTable = xTextTable.getAnchor().getStart();
        final XTextRange xEndTable = xTextTable.getAnchor().getEnd();

        // creates a view cursor from current cursor
        final XTextViewCursorSupplier xTextViewCursorSupplier = queryInterface(XTextViewCursorSupplier.class,xCurrentTextDocument.getCurrentController());
        final XTextViewCursor textViewCursor = xTextViewCursorSupplier.getViewCursor();

        final XLineCursor lineCursor = queryInterface(XLineCursor.class, textViewCursor);
        final XText xText = xCurrentTextDocument.getText();
        final XTextCursor xCursor = xText.createTextCursor();
        final XParagraphCursor xParagraphCursor = queryInterface(XParagraphCursor.class, xCursor);

        // moves the cursor to the line right before the table
        textViewCursor.gotoRange(xStartTable, false);

        // moves the cursor to the end of the line right before the table
        lineCursor.gotoEndOfLine(false);
        xCursor.gotoRange(textViewCursor.getStart(), false);

        // creates a new dynamic field in a new paragraph
        insertField(xCurrentTextDocument,xParagraphCursor.getStart(), "begin", "any");

        // moves the cursor to the line right after the table
        textViewCursor.gotoRange(xEndTable, false);
        textViewCursor.gotoEnd(false);

        // moves the cursor to the beginning of the line right after the table
        lineCursor.gotoStartOfLine(false);
        xCursor.gotoRange(textViewCursor.getEnd(), false);

        // creates a new dynamic field in a new paragraph
        insertField(xCurrentTextDocument,xParagraphCursor.getStart(), "end", "any");

        // moves the cursor to the position the user first selected
        textViewCursor.gotoRange(currentCursor, false);

    }

    static XTextField insertField(final XTextDocument xTextDocument, XTextRange xTextRange, String name, String value) throws Exception {
        final XTextFieldsSupplier xTextFieldsSupplier = queryInterface(XTextFieldsSupplier.class, xTextDocument);

        final XMultiServiceFactory xDocFactory = queryInterface(XMultiServiceFactory.class, xTextDocument);
        final Object field = xDocFactory.createInstanceWithArguments("com.sun.star.text.TextField.SetExpression", null);

        final XNameAccess xNamedFieldMasters = xTextFieldsSupplier.getTextFieldMasters();

        Object masterField;
        if (!xNamedFieldMasters.hasByName("com.sun.star.text.fieldmaster.SetExpression." + name)) {
            masterField = xDocFactory.createInstanceWithArguments("com.sun.star.text.fieldmaster.SetExpression", null);
        } else {
            masterField = xNamedFieldMasters.getByName("com.sun.star.text.fieldmaster.SetExpression." + name);
        }

        XPropertySet xMasterFieldPropSet = queryInterface(XPropertySet.class, masterField);
        xMasterFieldPropSet.setPropertyValue("Name", name);
        XDependentTextField xdtf = queryInterface(XDependentTextField.class, field);
        xdtf.attachTextFieldMaster(xMasterFieldPropSet);

        XPropertySet xPropSet = queryInterface(XPropertySet.class, field);
        xPropSet.setPropertyValue("Content", value);
        xPropSet.setPropertyValue("CurrentPresentation", value);
        xPropSet.setPropertyValue("IsVisible", true);
        xPropSet.setPropertyValue("IsInput", false);
        xPropSet.setPropertyValue("SubType", (short) 0);


        XTextViewCursorSupplier xTextViewCursorSupplier = queryInterface(XTextViewCursorSupplier.class, xTextDocument.getCurrentController());
        final XTextViewCursor textViewCursor = xTextViewCursorSupplier.getViewCursor();

        final XTextRange start = textViewCursor.getStart();
        final XTextRange end = textViewCursor.getEnd();
        textViewCursor.gotoRange(xTextRange, false);

        final XTextField xTextField = queryInterface(XTextField.class, field);
        textViewCursor.getText().insertTextContent(xTextRange, xTextField, false);

        textViewCursor.gotoRange(start, false);
        textViewCursor.gotoRange(end, true);

        return xTextField;
    }

}
