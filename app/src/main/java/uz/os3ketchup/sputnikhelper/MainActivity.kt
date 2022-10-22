package uz.os3ketchup.sputnikhelper

import android.annotation.SuppressLint
import android.app.Notification.Action
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.os.StrictMode
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.annotation.RequiresApi
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.FileProvider
import androidx.core.net.toUri
import org.apache.poi.wp.usermodel.HeaderFooterType
import org.apache.poi.xwpf.extractor.XWPFWordExtractor
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.security.Permission


class MainActivity : AppCompatActivity() {
    private lateinit var textView: TextView
    private lateinit var buttonView: Button




    @RequiresApi(Build.VERSION_CODES.O)
    @SuppressLint("SdCardPath")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        textView = findViewById(R.id.textView)
        buttonView = findViewById(R.id.buttonToShare)

        val targetDoc = createWordDoc()
        addParagraph(targetDoc)
        addTable(targetDoc)
        addHeaderAndFooter(targetDoc)
        saveOurDoc(targetDoc)
        readDoc()

        buttonView.setOnClickListener {
            val contentUri = saveOurDoc(targetDoc)
            val intent = Intent(Intent.ACTION_SEND)
            intent.type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            intent.putExtra(Intent.EXTRA_SUBJECT,"Subject here")
            intent.putExtra(Intent.EXTRA_STREAM,contentUri)
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            startActivity(Intent.createChooser(intent,"Share via"))


        }
    }


    private fun createWordDoc(): XWPFDocument {
        return XWPFDocument()
    }

    private fun addParagraph(targetDoc: XWPFDocument) {
        // creating a paragraph in our document and setting its alignment
        val paragraph1 = targetDoc.createParagraph()
        paragraph1.alignment = ParagraphAlignment.LEFT

        // creating a run for adding text
        val sentenceRun1 = paragraph1.createRun()

        //format the text
        sentenceRun1.isBold = true
        sentenceRun1.fontSize = 15
        sentenceRun1.fontFamily = "Comic Sans MS"
        sentenceRun1.setText("First sentence run starts here. It's such an honour too see you here :-)")
        //add a sentence break
        sentenceRun1.addBreak()

        //add another run
        val sentenceRun2 = paragraph1.createRun()
        sentenceRun2.fontSize = 12
        sentenceRun2.fontFamily = "Comic Sans MS"
        sentenceRun2.setText(
            "999Second sentence run starts here." +
                    " We love Apache POI. Lorem ipsum dolor sit amet," +
                    " consectetur adipiscing elit." +
                    " Sed lacinia dui consectetur euismod ultrices." +
                    " Aenean et enim pulvinar purus scelerisque dapibus." +
                    " Duis euismod lorem nec justo viverra ornare." +
                    " Aliquam est erat, mollis at iaculis eu, ultricies aliquet risus." +
                    " Proin lacinia ligula sed quam elementum, congue tincidunt lorem iaculis." +
                    " Nulla facilisi. Praesent faucibus metus eu nisi tincidunt rhoncus vitae et ligula." +
                    " Pellentesque quam dui, pellentesque vitae placerat eu, tempor ut lectus."
        )
        sentenceRun2.addBreak()
    }

    private fun addTable(targetDoc: XWPFDocument) {
        val ourTable = targetDoc.createTable()

        //Creating the first row and adding cell values
        val row1 = ourTable.getRow(0)
        row1.getCell(0).text = "Code"
        row1.addNewTableCell().text = "Item"

        //Creating the second row
        val row2 = ourTable.createRow()
        row2.getCell(0).text = "0345"
        row2.getCell(1).text = "Benz"

        //Creating the third row
        val row3 = ourTable.createRow()
        row3.getCell(0).text = "48542"
        row3.getCell(1).text = "Eng-Ed"
    }

    //adding a header and a footer
    private fun addHeaderAndFooter(targetDoc: XWPFDocument) {

        // initializing the header
        val docHeader = targetDoc.createHeader(HeaderFooterType.DEFAULT)

        // creating a run for the header. This is for setting the header text and styling
        val headerRun = docHeader.createParagraph().createRun()
        headerRun.setText("This is the header!")
        headerRun.fontFamily = "CopperPlate Gothic"
        headerRun.isBold = true
        headerRun.color = "00ff00"

        // initializing the footer
        val docFooter = targetDoc.createFooter(HeaderFooterType.DEFAULT)

        //creating a run for the footer. This sets the footer text and styling
        val footerRun = docFooter.createParagraph().createRun()
        footerRun.fontFamily = "Copperplate Gothic"
        footerRun.isBold = true
        footerRun.setText("This is the footer!")
    }

    //saving the word document
    private fun saveOurDoc(targetDoc: XWPFDocument):Uri {

        val ourAppFileDirectory = File(filesDir,"files")

        // Check whether it exits or not, and create one if it does not exist.
        if (!ourAppFileDirectory.exists()) {
            ourAppFileDirectory.mkdirs()

        }
        //create a word file called test.docx and save it to the file system
        val wordFile = File(ourAppFileDirectory,"myDoc.docx")
        var contentUri:Uri? = null
        try {
            val fileOut = FileOutputStream(wordFile)
            targetDoc.write(fileOut)
             contentUri = FileProvider.getUriForFile(this,"uz.os3ketchup.sputnikhelper.provider",wordFile)
            fileOut.close()
        } catch (e: IOException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        }
        return contentUri!!

    }

    // retrieving the document from the system
    private fun loadDoc(): File? {
        val ourDirectory = File(filesDir,"files")
        ourDirectory.let {
            //Check if file exists or not
            if (it.exists()) {
                //check the file in the directory called myDoc.docx
                // return the file


                return File(ourDirectory, "myDoc.docx")
            }
        }
        return null
    }

    // reading the document's text
    private fun readDoc() {
        loadDoc().let {
            try {
                //Reading it as stream
                val docStream = FileInputStream(it)
                val targetDoc = XWPFDocument(docStream)

                //creating a constructor object for extracting text from the word document
                val wordExtractor = XWPFWordExtractor(targetDoc)
                val docText = wordExtractor.text
                //displaying the text read from the document
                textView.text = docText
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }


}