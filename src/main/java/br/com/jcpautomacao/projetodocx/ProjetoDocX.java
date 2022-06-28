/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package br.com.jcpautomacao.projetodocx;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import javax.lang.model.element.Modifier;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakClear;
import static org.apache.poi.xwpf.usermodel.BreakClear.ALL;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

/**
 * A simple WOrdprocessingML document created by POI XWPF API
 */
public class ProjetoDocX {

    public static void main(String[] args) throws Exception {
        /**
         *
         * @throws IOException
         */
        paragrafos();
    }

    private static void paragrafos() {

        FileInputStream fs = null;
        XWPFDocument doc = null;
        String text = null;

        try {
            fs = new FileInputStream("modelo_laudo_poi.docx");

        } catch (FileNotFoundException e) {

            System.out.print(e.getMessage());
        }

        try {
            doc = new XWPFDocument(fs);

        } catch (IOException ex) {
            System.out.print(ex.getMessage());

        }
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                //List<XWPFRun> runs = paragraph.getRuns();
                if (run != null && run.getText(0) != null) {
                    text = run.getText(0);
                    
                    if (text.contains("<Processo>")) {
                        text = text.replace("<Processo>", "1234567-12.1234.1.12.2222");                        
                        run.setText(text, 0);
                        
                    } else if (text.contains("#Reclamante")) {
                        text = text.replace("#Reclamante", "Antonio Carlos Fonseca");
                        run.setText(text, 0);
                        
                    } else if (text.contains("#Reclamada")) {
                        text = text.replace("#Reclamada", "Benedito Fonseca");
                        run.setText(text, 0);
                        
                    }
                }
                System.out.print(run + "\n");
                System.out.print(text);
            }
        }

        try ( FileOutputStream out = new FileOutputStream("modelo_laudo_poi_alt-v2.docx")) {
            doc.write(out);

        } catch (Exception e) {
            System.out.print("Error ao abrir file outputstream");

        }
    }
}
