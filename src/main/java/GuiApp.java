import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * wordAnyWhere 图形界面版：选择 Word 模板与 Excel 数据表，把表格中每一行
 * 按模板批量生成文档，并按"标识列"的值分文件夹输出。
 *
 * 界面基于 Swing，编译目标为 Java 8，可在 Windows XP ~ Windows 11 上运行
 * （XP 需已安装 JRE 8；Win7 及以上可使用自带的 jdk-15.0.2）。
 */
public class GuiApp {

    private JFrame frame;
    private JTextField templateField;
    private JTextField excelField;
    private JTextField outputField;
    private JComboBox<String> idColumnBox;
    private JTextArea logArea;
    private JButton checkButton;
    private JButton runButton;

    private final List<File> templates = new ArrayList<>();
    private volatile File excelFile;
    private volatile ExcelData excelData;
    private boolean excelLoading;
    private boolean running;
    private int loadGeneration;
    private File lastDir;

    /** Excel 读取结果：跨工作表汇总的表头 + 行数据 */
    static class ExcelData {
        final List<String> headers = new ArrayList<>();
        final List<RowData> rows = new ArrayList<>();
    }

    /** 一行数据及其显示标签（如 "Sheet1 第3行"） */
    static class RowData {
        final String label;
        final Map<String, String> data = new LinkedHashMap<>();

        RowData(String label) {
            this.label = label;
        }
    }

    /** 批量执行的日志出口，便于界面与测试共用 */
    interface LogSink {
        void log(String line);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    javax.swing.UIManager.setLookAndFeel(javax.swing.UIManager.getSystemLookAndFeelClassName());
                } catch (Exception ignored) {
                }
                new GuiApp().frame.setVisible(true);
            }
        });
    }

    GuiApp() {
        frame = new JFrame("Word 批量生成工具 (wordAnyWhere GUI)");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setPreferredSize(new Dimension(720, 560));

        JPanel root = new JPanel(new GridBagLayout());
        root.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(4, 4, 4, 4);
        gc.fill = GridBagConstraints.HORIZONTAL;

        gc.gridx = 0; gc.gridy = 0; gc.weightx = 0;
        root.add(new JLabel("模板文件(*.docx，可多选):"), gc);
        templateField = new JTextField();
        templateField.setEditable(false);
        gc.gridx = 1; gc.weightx = 1;
        root.add(templateField, gc);
        JButton tplButton = new JButton("浏览...");
        gc.gridx = 2; gc.weightx = 0;
        root.add(tplButton, gc);
        tplButton.addActionListener(e -> chooseTemplates());

        gc.gridx = 0; gc.gridy = 1;
        root.add(new JLabel("Excel数据表(*.xlsx):"), gc);
        excelField = new JTextField();
        excelField.setEditable(false);
        gc.gridx = 1; gc.weightx = 1;
        root.add(excelField, gc);
        JButton excelButton = new JButton("浏览...");
        gc.gridx = 2; gc.weightx = 0;
        root.add(excelButton, gc);
        excelButton.addActionListener(e -> chooseExcel());

        gc.gridx = 0; gc.gridy = 2;
        root.add(new JLabel("输出目录:"), gc);
        outputField = new JTextField();
        gc.gridx = 1; gc.weightx = 1;
        root.add(outputField, gc);
        JButton outButton = new JButton("浏览...");
        gc.gridx = 2; gc.weightx = 0;
        root.add(outButton, gc);
        outButton.addActionListener(e -> chooseOutputDir());

        gc.gridx = 0; gc.gridy = 3;
        root.add(new JLabel("目录命名列(标识列):"), gc);
        idColumnBox = new JComboBox<>();
        gc.gridx = 1; gc.weightx = 1;
        root.add(idColumnBox, gc);
        checkButton = new JButton("检查模板");
        gc.gridx = 2; gc.weightx = 0;
        root.add(checkButton, gc);
        checkButton.addActionListener(e -> checkTemplates());

        runButton = new JButton("开始执行");
        gc.gridx = 1; gc.gridy = 4; gc.weightx = 1;
        gc.anchor = GridBagConstraints.LINE_START;
        root.add(runButton, gc);
        gc.anchor = GridBagConstraints.CENTER;
        runButton.addActionListener(e -> startRun());

        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setLineWrap(true);
        logArea.setWrapStyleWord(true);
        JScrollPane scroll = new JScrollPane(logArea);
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        scroll.setPreferredSize(new Dimension(690, 300));
        gc.gridx = 0; gc.gridy = 5; gc.gridwidth = 3;
        gc.weighty = 1;
        gc.fill = GridBagConstraints.BOTH;
        root.add(scroll, gc);

        frame.setContentPane(root);
        outputField.setText(System.getProperty("user.dir") + File.separator + "output");
        frame.pack();
        frame.setLocationRelativeTo(null);
        appendLog("就绪。请依次选择模板、Excel 数据表和输出目录，");
        appendLog("点击【检查模板】核对占位符与 Excel 列名，然后【开始执行】。");
    }

    // ------------------------------------------------------------------
    // 界面动作
    // ------------------------------------------------------------------

    private void chooseTemplates() {
        JFileChooser fc = newChooser("Word 模板 (*.docx)", "docx");
        fc.setMultiSelectionEnabled(true);
        if (fc.showOpenDialog(frame) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        templates.clear();
        Set<String> seen = new HashSet<>();
        for (File f : fc.getSelectedFiles()) {
            String key = f.getAbsolutePath().toLowerCase();
            if (seen.add(key)) {
                templates.add(f);
            }
        }
        templateField.setText(describeTemplates());
        log("已选择模板 " + templates.size() + " 个:");
        for (File f : templates) {
            log("  " + f.getAbsolutePath());
        }
    }

    private void chooseExcel() {
        JFileChooser fc = newChooser("Excel 数据表 (*.xlsx; *.xlsm)", "xlsx", "xlsm");
        fc.setMultiSelectionEnabled(false);
        if (fc.showOpenDialog(frame) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        excelFile = fc.getSelectedFile();
        excelField.setText(excelFile.getAbsolutePath());
        loadExcelAsync();
    }

    private void chooseOutputDir() {
        JFileChooser fc = newChooser(null);
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        fc.setMultiSelectionEnabled(false);
        if (fc.showOpenDialog(frame) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        outputField.setText(fc.getSelectedFile().getAbsolutePath());
    }

    private JFileChooser newChooser(String description, String... extensions) {
        JFileChooser fc = new JFileChooser(lastDir);
        if (description != null) {
            fc.setFileFilter(new FileNameExtensionFilter(description, extensions));
        }
        return fc;
    }

    private String describeTemplates() {
        if (templates.size() == 1) {
            return templates.get(0).getAbsolutePath();
        }
        StringBuilder sb = new StringBuilder("共 ").append(templates.size()).append(" 个: ");
        int used = 0;
        for (File f : templates) {
            String name = f.getName();
            if (used > 0 && used + name.length() < 60) {
                sb.append("; ");
            } else if (used > 0) {
                sb.append(" ...");
                break;
            }
            sb.append(name);
            used += name.length() + 2;
        }
        return sb.toString();
    }

    private void loadExcelAsync() {
        final File file = excelFile;
        final int generation = ++loadGeneration;
        excelData = null;
        excelLoading = true;
        updateButtons();
        log("正在读取 Excel: " + file.getAbsolutePath());
        new SwingWorker<ExcelData, Void>() {
            @Override
            protected ExcelData doInBackground() throws IOException {
                return readExcel(file);
            }

            @Override
            protected void done() {
                if (generation != loadGeneration) {
                    return; // 用户又选了新的 Excel，本次结果作废
                }
                excelLoading = false;
                try {
                    ExcelData d = get();
                    excelData = d;
                    idColumnBox.removeAllItems();
                    int preferred = 0;
                    for (int i = 0; i < d.headers.size(); i++) {
                        idColumnBox.addItem(d.headers.get(i));
                        if ("店铺名".equals(d.headers.get(i))) {
                            preferred = i;
                        }
                    }
                    if (d.headers.size() > 0) {
                        idColumnBox.setSelectedIndex(preferred);
                    }
                    log("√ Excel 读取完成: " + d.headers.size() + " 个字段, " + d.rows.size() + " 行数据");
                    log("  字段: " + String.join(", ", d.headers));
                } catch (Exception e) {
                    excelData = null;
                    log("× Excel 读取失败: " + e);
                    showError("Excel 读取失败:\n" + e);
                }
                updateButtons();
            }
        }.execute();
    }

    private void checkTemplates() {
        if (templates.isEmpty()) {
            showError("请先选择模板文件。");
            return;
        }
        if (excelData == null) {
            showError("请先选择 Excel 数据表并等待读取完成。");
            return;
        }
        checkButton.setEnabled(false);
        log("—— 模板检查 ——");
        new SwingWorker<String, Void>() {
            @Override
            protected String doInBackground() throws IOException {
                return buildCheckReport();
            }

            @Override
            protected void done() {
                try {
                    log(get());
                } catch (Exception e) {
                    log("× 检查失败: " + e);
                } finally {
                    updateButtons();
                }
            }
        }.execute();
    }

    private void startRun() {
        if (!validateForRun()) {
            return;
        }
        final String idCol = (String) idColumnBox.getSelectedItem();
        final File outBase = new File(outputField.getText().trim());
        setRunning(true);
        log("—— 开始执行 ——");
        new SwingWorker<int[], String>() {
            @Override
            protected int[] doInBackground() {
                return runBatch(outBase, idCol, new LogSink() {
                    @Override
                    public void log(String line) {
                        publish(line);
                    }
                });
            }

            @Override
            protected void process(List<String> chunks) {
                for (String line : chunks) {
                    appendLog(line);
                }
            }

            @Override
            protected void done() {
                try {
                    int[] c = get();
                    log("—— 完成 ——");
                    log("成功生成 " + c[0] + " 个文件，失败 " + c[1] + " 个，跳过 " + c[2] + " 行。");
                    log("输出目录: " + outBase.getAbsolutePath());
                } catch (Exception e) {
                    log("× 执行中断: " + e);
                } finally {
                    setRunning(false);
                }
            }
        }.execute();
    }

    private boolean validateForRun() {
        if (templates.isEmpty()) {
            showError("请先选择模板文件。");
            return false;
        }
        if (excelData == null) {
            showError("请先选择 Excel 数据表并等待读取完成。");
            return false;
        }
        if (outputField.getText().trim().isEmpty()) {
            showError("请设置输出目录。");
            return false;
        }
        if (idColumnBox.getSelectedItem() == null) {
            showError("请选择目录命名列(标识列)。");
            return false;
        }
        return true;
    }

    private void setRunning(boolean value) {
        running = value;
        runButton.setEnabled(!running && !excelLoading);
        runButton.setText(running ? "执行中..." : "开始执行");
        checkButton.setEnabled(!running && !excelLoading);
    }

    private void updateButtons() {
        setRunning(running);
    }

    private void showError(String message) {
        JOptionPane.showMessageDialog(frame, message, "提示", JOptionPane.WARNING_MESSAGE);
    }

    private void log(String line) {
        if (SwingUtilities.isEventDispatchThread()) {
            appendLog(line);
        } else {
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    appendLog(line);
                }
            });
        }
    }

    private void appendLog(String line) {
        logArea.append(line);
        logArea.append("\n");
        logArea.setCaretPosition(logArea.getDocument().getLength());
    }

    // ------------------------------------------------------------------
    // 核心逻辑（与界面解耦，便于单独测试）
    // ------------------------------------------------------------------

    /** 读取 Excel：所有工作表，第一行为表头，其余每行为一条数据 */
    static ExcelData readExcel(File file) throws IOException {
        ExcelData data = new ExcelData();
        Workbook wb = WorkbookFactory.create(file);
        try {
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                Sheet sheet = wb.getSheetAt(s);
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) {
                    continue;
                }
                List<Integer> colIndexes = new ArrayList<>();
                Map<Integer, String> colNames = new LinkedHashMap<>();
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    org.apache.poi.ss.usermodel.Cell cell = headerRow.getCell(c);
                    String header = cell == null ? "" : fmt.formatCellValue(cell, eval).trim();
                    if (header.isEmpty()) {
                        continue; // 表头为空的列不参与替换
                    }
                    colIndexes.add(c);
                    colNames.put(c, header);
                    if (!data.headers.contains(header)) {
                        data.headers.add(header);
                    }
                }
                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    RowData rd = new RowData(sheet.getSheetName() + " 第" + (r + 1) + "行");
                    boolean hasValue = false;
                    for (int c : colIndexes) {
                        org.apache.poi.ss.usermodel.Cell cell = row == null ? null : row.getCell(c);
                        String value = cell == null ? "" : fmt.formatCellValue(cell, eval);
                        rd.data.put(colNames.get(c), value);
                        if (!value.trim().isEmpty()) {
                            hasValue = true;
                        }
                    }
                    if (hasValue) {
                        data.rows.add(rd); // 整行为空的数据不生成文档
                    }
                }
            }
        } finally {
            wb.close();
        }
        return data;
    }

    /** 提取模板中全部占位符名（含正文、表格、页眉页脚），保持出现顺序 */
    static LinkedHashSet<String> extractTags(File tpl) throws IOException {
        XWPFTemplate t = XWPFTemplate.compile(tpl);
        try {
            LinkedHashSet<String> tags = new LinkedHashSet<>();
            for (MetaTemplate mt : t.getElementTemplates()) {
                if (mt instanceof ElementTemplate) {
                    String name = ((ElementTemplate) mt).getTagName();
                    if (name != null && !name.isEmpty()) {
                        tags.add(name);
                    }
                }
            }
            return tags;
        } finally {
            try {
                t.close();
            } catch (IOException ignored) {
            }
        }
    }

    /** 生成"占位符 = Excel列"对照报告 */
    String buildCheckReport() throws IOException {
        ExcelData d = excelData;
        StringBuilder sb = new StringBuilder();
        Set<String> used = new LinkedHashSet<>();
        int total = 0;
        int missing = 0;
        for (File tpl : templates) {
            sb.append("模板[").append(tpl.getName()).append("]");
            LinkedHashSet<String> tags;
            try {
                tags = extractTags(tpl);
            } catch (Exception e) {
                sb.append("\n  × 读取失败: ").append(e).append("\n");
                continue;
            }
            if (tags.isEmpty()) {
                sb.append("\n  提示: 未发现占位符，将原样复制");
            }
            for (String tag : tags) {
                total++;
                used.add(tag);
                if (d.headers.contains(tag)) {
                    sb.append("\n  √ {{").append(tag).append("}} = Excel列[").append(tag).append("]");
                } else {
                    missing++;
                    sb.append("\n  × {{").append(tag).append("}} —— Excel 中不存在该列，输出时将被删除");
                }
            }
            sb.append("\n");
        }
        List<String> unused = new ArrayList<>();
        for (String h : d.headers) {
            if (!used.contains(h)) {
                unused.add(h);
            }
        }
        if (!unused.isEmpty()) {
            sb.append("提示: 以下 Excel 列未被任何模板使用: ").append(String.join(", ", unused)).append("\n");
        }
        sb.append("汇总: 占位符共 ").append(total).append(" 个，匹配 ").append(total - missing)
                .append(" 个，缺失 ").append(missing).append(" 个");
        if (excelData != null && !excelData.headers.isEmpty()) {
            String idCol = (String) idColumnBox.getSelectedItem();
            if (idCol != null) {
                sb.append(excelData.headers.contains(idCol) ? "\n标识列[" + idCol + "] √ 存在"
                        : "\n标识列[" + idCol + "] × 不存在，请重新选择");
            }
        }
        return sb.toString();
    }

    /**
     * 批量渲染。返回统计: {成功文件数, 失败文件数, 跳过行数}。
     * 单个模板或单行失败只记录日志并继续，不中断整批。
     */
    int[] runBatch(File outBase, String idCol, LogSink sink) {
        int[] counts = new int[3];
        ExcelData data = excelData;
        List<File> tplList = new ArrayList<>(templates);
        if (!outBase.exists() && !outBase.mkdirs()) {
            sink.log("× 无法创建输出目录: " + outBase.getAbsolutePath());
            return counts;
        }
        Set<String> madeDirs = new HashSet<>();
        Set<String> writtenFiles = new HashSet<>();
        for (RowData rd : data.rows) {
            String id = rd.data.get(idCol);
            id = id == null ? "" : id.trim();
            if (id.isEmpty()) {
                counts[2]++;
                sink.log("跳过 " + rd.label + "：标识列[" + idCol + "]为空");
                continue;
            }
            String folder = sanitizeName(id);
            File dir = new File(outBase, folder);
            if (!madeDirs.contains(folder) && !dir.exists()) {
                if (!dir.mkdirs()) {
                    counts[2]++;
                    sink.log("跳过 " + rd.label + "：创建目录失败 " + dir.getAbsolutePath());
                    continue;
                }
                madeDirs.add(folder);
            }
            sink.log(rd.label + " → " + folder + "/");
            for (File tpl : tplList) {
                String newName = sanitizeName(replaceName(tpl.getName(), rd.data));
                File dest = new File(dir, newName);
                if (!writtenFiles.add(dest.getAbsolutePath().toLowerCase())) {
                    sink.log("  提示: 同名文件将被覆盖 - " + newName);
                }
                try {
                    render(tpl, rd.data, dest);
                    counts[0]++;
                    sink.log("  √ " + newName);
                } catch (Exception ex) {
                    counts[1]++;
                    sink.log("  × " + tpl.getName() + " 生成失败: " + ex.getMessage());
                }
            }
        }
        return counts;
    }

    /** 用一行数据替换文件名中的 {{字段名}} */
    static String replaceName(String name, Map<String, String> data) {
        for (Map.Entry<String, String> entry : data.entrySet()) {
            name = name.replace("{{" + entry.getKey() + "}}", entry.getValue());
        }
        return name;
    }

    /** 渲染单个模板到目标文件 */
    static void render(File tpl, Map<String, String> data, File dest) throws IOException {
        XWPFTemplate t = XWPFTemplate.compile(tpl);
        FileOutputStream fos = null;
        try {
            t.render(data);
            fos = new FileOutputStream(dest);
            t.write(fos);
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException ignored) {
                }
            }
            try {
                t.close();
            } catch (IOException ignored) {
            }
        }
    }

    /** 去除 Windows 文件名/目录名中的非法字符 */
    static String sanitizeName(String s) {
        String r = s.replaceAll("[\\\\/:*?\"<>|\\r\\n\\t]", "_").trim();
        while (r.endsWith(".") || r.endsWith(" ")) {
            r = r.substring(0, r.length() - 1);
        }
        if (r.length() > 100) {
            r = r.substring(0, 100);
        }
        return r.isEmpty() ? "_" : r;
    }
}
