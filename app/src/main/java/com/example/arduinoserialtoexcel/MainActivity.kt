package com.example.arduinoserialtoexcel

import android.app.Activity
import android.content.Context
import android.content.Intent
import android.hardware.usb.UsbManager
import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.view.View
import android.widget.*
import androidx.activity.result.ActivityResultLauncher
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.Job
import kotlinx.coroutines.cancel
import kotlinx.coroutines.delay
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.OutputStream
import java.nio.charset.Charset
import com.hoho.android.usbserial.driver.UsbSerialProber
import com.hoho.android.usbserial.util.SerialInputOutputManager
import com.hoho.android.usbserial.driver.UsbSerialPort
import com.hoho.android.usbserial.driver.UsbSerialDriver
import java.util.concurrent.Executors

class MainActivity : AppCompatActivity() {
    companion object {
        private const val TAG = "MainActivity"
        private const val TRIGGER_TEXT = "File created and data written."
        private const val END_TEXT = "--- END OF FILE ---"
    }

    private lateinit var usbManager: UsbManager
    private var driver: UsbSerialDriver? = null
    private var port: UsbSerialPort? = null
    private var ioManager: SerialInputOutputManager? = null

    private lateinit var btnDiscover: Button
    private lateinit var btnConnect: Button
    private lateinit var btnWait: Button
    private lateinit var btnRead: Button
    private lateinit var btnExport: Button
    private lateinit var btnAuto: Button
    private lateinit var statusTv: TextView
    private lateinit var logTv: TextView
    private lateinit var listView: ListView

    private val coroutineScope = CoroutineScope(Dispatchers.Main + Job())
    private val collectedLines = mutableListOf<String>()
    private val parsedRows = mutableListOf<ParsedRow>()
    private lateinit var createDocumentLauncher: ActivityResultLauncher<String>

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(createLayout())

        usbManager = getSystemService(Context.USB_SERVICE) as UsbManager

        createDocumentLauncher = registerForActivityResult(ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) { uri ->
            uri?.let { uri ->
                coroutineScope.launch {
                    saveExcelToUri(uri)
                }
            }
        }

        btnDiscover.setOnClickListener { discoverDevices() }
        btnConnect.setOnClickListener { connectToSelectedDevice() }
        btnWait.setOnClickListener { coroutineScope.launch { waitForTrigger() } }
        btnRead.setOnClickListener { coroutineScope.launch { readLogFromArduino() } }
        btnExport.setOnClickListener { createDocumentLauncher.launch("arduino_log.xlsx") }
        btnAuto.setOnClickListener { coroutineScope.launch { autoFlow() } }

        updateUiState()
    }

    override fun onDestroy() {
        super.onDestroy()
        stopIoManager()
        coroutineScope.cancel()
    }

    private fun createLayout(): View {
        val root = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            setPadding(16, 16, 16, 16)
        }

        val title = TextView(this).apply { text = "Arduino Serial → Excel"; textSize = 20f }
        root.addView(title)

        statusTv = TextView(this).apply { text = "Status: idle" }
        root.addView(statusTv)

        val controls = LinearLayout(this).apply { orientation = LinearLayout.HORIZONTAL }
        btnDiscover = Button(this).apply { text = "Discover" }
        btnConnect = Button(this).apply { text = "Connect" }
        btnWait = Button(this).apply { text = "Wait for trigger" }
        btnRead = Button(this).apply { text = "Send 'l' & Read" }
        btnExport = Button(this).apply { text = "Export Excel" }
        btnAuto = Button(this).apply { text = "Auto" }

        controls.addView(btnDiscover)
        controls.addView(btnConnect)
        controls.addView(btnWait)
        root.addView(controls)

        val controls2 = LinearLayout(this).apply { orientation = LinearLayout.HORIZONTAL }
        controls2.addView(btnRead)
        controls2.addView(btnExport)
        controls2.addView(btnAuto)
        root.addView(controls2)

        logTv = TextView(this).apply {
            text = "\nLog:\n"
            setLines(10)
            isVerticalScrollBarEnabled = true
        }
        val scroll = ScrollView(this)
        scroll.addView(logTv)
        root.addView(scroll, LinearLayout.LayoutParams(LinearLayout.LayoutParams.MATCH_PARENT, 400))

        listView = ListView(this)
        root.addView(listView, LinearLayout.LayoutParams(LinearLayout.LayoutParams.MATCH_PARENT, LinearLayout.LayoutParams.WRAP_CONTENT))

        return root
    }

    private fun appendLog(s: String) {
        runOnUiThread { logTv.append(s + "\n") }
    }

    private fun updateUiState() {
        runOnUiThread {
            btnConnect.isEnabled = driver != null && port == null
            btnWait.isEnabled = port != null
            btnRead.isEnabled = port != null
            btnExport.isEnabled = parsedRows.isNotEmpty()
            btnAuto.isEnabled = port != null
        }
    }

    private fun discoverDevices() {
        appendLog("Discovering USB devices...")
        val availableDrivers = UsbSerialProber.getDefaultProber().findAllDrivers(usbManager)
        if (availableDrivers.isEmpty()) {
            appendLog("No serial drivers found. Make sure device is connected and OTG cable is used.")
            showSimpleDialog("No devices", "No serial devices found. Is the Arduino connected with an OTG cable?")
            return
        }

        if (availableDrivers.size == 1) {
            driver = availableDrivers[0]
            appendLog("Found device: " + driver!!.device.deviceName)
            showSimpleDialog("Device found", "Found: ${driver!!.device.deviceName}")
        } else {
            val names = availableDrivers.map { it.device.deviceName }
            val arr = names.toTypedArray()
            AlertDialog.Builder(this)
                .setTitle("Select device")
                .setItems(arr) { _, which ->
                    driver = availableDrivers[which]
                    appendLog("Selected: ${driver!!.device.deviceName}")
                }
                .show()
        }
        updateUiState()
    }

    private fun connectToSelectedDevice() {
        val selDriver = driver ?: run {
            showSimpleDialog("No device", "Please discover a device first.")
            return
        }
        val device = selDriver.device
        usbManager.requestPermission(device, null)

        coroutineScope.launch {
            delay(500)
            try {
                val connection = usbManager.openDevice(device)
                val ports = selDriver.ports
                if (ports.isEmpty()) throw Exception("No ports")
                port = ports[0]
                port!!.open(connection)
                port!!.setParameters(9600, 8, UsbSerialPort.STOPBITS_1, UsbSerialPort.PARITY_NONE)
                startIoManager()
                appendLog("Connected to ${device.deviceName}")
                statusTv.text = "Connected: ${device.deviceName}"
                updateUiState()
            } catch (e: Exception) {
                appendLog("Failed to open port: ${e.message}")
            }
        }
    }

    private fun startIoManager() {
        stopIoManager()
        val executor = Executors.newSingleThreadExecutor()
        ioManager = SerialInputOutputManager(port, object : SerialInputOutputManager.Listener {
            override fun onNewData(data: ByteArray) {
                val s = String(data, Charset.forName("UTF-8"))
                handleIncomingText(s)
            }

            override fun onRunError(e: Exception) {
                appendLog("IO Manager error: ${e.message}")
            }
        })
        executor.submit(ioManager)
    }

    private fun stopIoManager() {
        try { ioManager?.stop() } catch (e: Exception) {}
        ioManager = null
    }

    private fun handleIncomingText(chunk: String) {
        val parts = chunk.replace("\r", "").split('\n')
        for (p in parts) {
            val line = p.trim()
            if (line.isNotEmpty()) {
                appendLog(line)
                synchronized(collectedLines) {
                    collectedLines.add(line)
                }
            }
        }
    }

    private suspend fun waitForTrigger() {
        appendLog("Waiting for trigger: $TRIGGER_TEXT")
        withContext(Dispatchers.Default) {
            while (true) {
                synchronized(collectedLines) {
                    if (collectedLines.any { it.contains(TRIGGER_TEXT) }) return@withContext
                }
                delay(100)
            }
        }
        appendLog("Trigger detected")
    }

    private suspend fun readLogFromArduino() {
        withContext(Dispatchers.IO) {
            appendLog("Sending 'l' to Arduino...")
            try {
                port?.write(ByteArray(1) { 'l'.code.toByte() }, 1000)
            } catch (e: Exception) {
                appendLog("Write failed: ${e.message}")
                return@withContext
            }

            synchronized(collectedLines) {
                collectedLines.clear()
            }

            while (true) {
                var found = false
                synchronized(collectedLines) {
                    if (collectedLines.any { it.contains(END_TEXT) }) found = true
                }
                if (found) break
                delay(100)
            }

            val copy: List<String>
            synchronized(collectedLines) { copy = ArrayList(collectedLines) }

            appendLog("End of file detected; parsing ${copy.size} lines")
            parseLines(copy)
            withContext(Dispatchers.Main) { renderPreview() }
        }
    }

    private data class ParsedRow(val ts: String, val temp: Double, val hum: Double)

    private fun parseLines(lines: List<String>) {
        parsedRows.clear()
        for (line in lines) {
            if (line.contains("Temp") && line.contains("Humidity")) {
                try {
                    val parts = line.split(',')
                    val ts = parts[0].trim()
                    val tempPart = parts[1].split(':')[1].replace("°C", "").replace("C", "").trim()
                    val temp = tempPart.toDouble()
                    val humPart = parts[2].split(':')[1].replace("%", "").trim()
                    val hum = humPart.toDouble()
                    parsedRows.add(ParsedRow(ts, temp, hum))
                } catch (e: Exception) {
                    appendLog("Skipping line due to parse error: $line (${e.message})")
                }
            }
        }
        appendLog("Parsed ${parsedRows.size} rows")
        updateUiState()
    }

    private fun renderPreview() {
        val list = parsedRows.mapIndexed { i, r -> "${i+1}. ${r.ts} — ${r.temp}°C — ${r.hum}%" }
        val adapter = ArrayAdapter(this, android.R.layout.simple_list_item_1, list)
        listView.adapter = adapter
    }

    private suspend fun saveExcelToUri(uri: Uri) {
        withContext(Dispatchers.IO) {
            try {
                val out = contentResolver.openOutputStream(uri)
                out?.use { generateExcel(it) }
                appendLog("Excel saved to chosen location")
            } catch (e: Exception) {
                appendLog("Failed to save Excel: ${e.message}")
            }
        }
    }

    private fun generateExcel(output: OutputStream) {
        val wb = XSSFWorkbook()
        val sheet = wb.createSheet("Log")
        val header = sheet.createRow(0)
        val headers = arrayOf("#", "Timestamp", "Temp (°C)", "Humidity (%)", "Utah", "Cumulative Utah")
        for ((i, h) in headers.withIndex()) {
            val c = header.createCell(i)
            c.setCellValue(h)
        }

        for ((i, row) in parsedRows.withIndex()) {
            val r = sheet.createRow(i + 1)
            r.createCell(0).setCellValue((i + 1).toDouble())
            r.createCell(1).setCellValue(row.ts)
            r.createCell(2).setCellValue(row.temp)
            r.createCell(3).setCellValue(row.hum)

            val excelRowNum = i + 2
            val tempCellRef = "C$excelRowNum"
            val utahFormula = "IF(${tempCellRef}=\"\",0,IF(${tempCellRef}<1.111,0,IF(${tempCellRef}<2.222,0.5,IF(${tempCellRef}<8.889,1,IF(${tempCellRef}<12.222,0.5,IF(${tempCellRef}<15.556,0,IF(${tempCellRef}<18.333,-0.5,-1)))))))"
            r.createCell(4).cellFormula = utahFormula
            r.createCell(5).cellFormula = "SUM(E$2:E$excelRowNum)"
        }

        sheet.setColumnWidth(0, 5 * 256)
        sheet.setColumnWidth(1, 20 * 256)
        sheet.setColumnWidth(2, 12 * 256)
        sheet.setColumnWidth(3, 12 * 256)
        sheet.setColumnWidth(4, 12 * 256)
        sheet.setColumnWidth(5, 16 * 256)

        wb.write(output)
        wb.close()
    }

    private fun autoFlow() {
        coroutineScope.launch {
            try {
                appendLog("Auto: wait -> read -> export")
                waitForTrigger()
                readLogFromArduino()
                createDocumentLauncher.launch("arduino_log.xlsx")
            } catch (e: Exception) {
                appendLog("Auto failed: ${e.message}")
            }
        }
    }

    private fun showSimpleDialog(title: String, message: String) {
        runOnUiThread {
            AlertDialog.Builder(this).setTitle(title).setMessage(message).setPositiveButton("OK", null).show()
        }
    }
}
