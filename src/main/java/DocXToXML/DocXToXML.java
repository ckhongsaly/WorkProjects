
package main.java.DocXToXML;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

//https://www.codota.com/code/java/classes/org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun

//https://github.com/Sayi/poi-tl/blob/master/src/main/java/com/deepoove/poi/XWPFParagraphWrapper.java#L49

//https://stackoverflow.com/questions/35088893/apache-poi-remove-cthyperlink-low-level-code/35101440#35101440

//https://www.w3.org/standards/xml/core

//https://docs.moodle.org/38/en/Moodle_XML_format

public class DocXToXML {

	private static XWPFHyperlink tempLink;

	/**
	 * Example of multichoice 
	 * <question type="multichoice"> <answer fraction="100">
	 * <text>The correct answer</text> <feedback><text>Correct!</text></feedback>
	 * </answer> <answer fraction="0"> <text>A distractor</text>
	 * <feedback><text>Ooops!</text></feedback> </answer> <answer fraction="0">
	 * <text>Another distractor</text> <feedback><text>Ooops!</text></feedback>
	 * </answer> <shuffleanswers>1</shuffleanswers> <single>true</single>
	 * <answernumbering>abc</answernumbering> </question>
	 */

	// For multiple choice questions
	private static String question_multichoice = "<question type=\"multichoice\">\n";
	private static String shuffle_answer = "<shuffleanswers>1</shuffleanswers>";
	private static String single = "<single>true</single>";
	private static String answer_numbering = "<answernumbering>abc</answernumbering>";

	// For all question types
	private static String end_of_question = "</question>";
	private static String correct = "<answer fraction=\"100\">";
	private static String incorrect = "<answer fraction=\"0\">";
	private static String start_of_question_text = "<text>";
	private static String end_of_question_text = "</text>";
	private static String end_of_answer = "</answer>";
	private static String feedback_start = "<feedback><text>";
	private static String feedback_end = "</text></feedback>";

	/**
	 * Description: Determine whether the first character in string is an integer
	 * 
	 * @param str
	 * @return return if first character is an integer
	 */
	public static boolean isInteger(String str) {
		return Character.isDigit(str.charAt(0));
	}

	/**
	 * Description: Determine whether the first character in string is the feedback
	 * 
	 * @param str
	 * @return return whether first character is ~
	 */
	public static boolean isFeedback(String str) {
		return str.charAt(0) == '~';
	}

	/**
	 * Description: Determine whether the first character in string is the answer
	 * 
	 * @param str
	 * @return return whether first character is *
	 */
	public static boolean isAnswer(String str) {
		return str.charAt(0) == '*';
	}

	/**
	 * Description: Determine whether the first character in string is the answer
	 * 
	 * @param str
	 * @return return whether first character is *
	 */
	public static boolean is(String str) {
		return str.charAt(0) == '*';
	}

	/**
	 * Description: Locate questions in path and return total number of questions
	 * 
	 * @param path location of file
	 * @return return total text found in link for docx
	 */
	public static int find_Question(String path) {

		System.out.println("Initalize find_Question..");

		int questionCounter = 0;
		int debugCounter = 0;

		try {
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));

			// Find and replace in paragraph
			List<XWPFParagraph> paragraphList = document.getParagraphs();

			for (XWPFParagraph paragraph : paragraphList) {
				List<XWPFRun> runs = paragraph.getRuns();

				if (runs != null) {

					for (XWPFRun r : runs) {

						// System.out.println(debugCounter++ + ". "+ r.toString());

						String tempString = r.toString();
						if (isInteger(tempString)) {
							questionCounter++;
							System.out.println(questionCounter + ". " + r.toString());
						}
					}
				}
			}

			// close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
		}

		catch (Exception ex) {
			ex.printStackTrace();
		}

		return questionCounter;

	}// end of find_Questions

	/**
	 * Description: Locate questions in path and return total number of questions
	 * 
	 * @param path location of file
	 * @return return total text found in link for docx
	 */
	public static int find_Feedback(String path) {

		System.out.println("Initalize find_Feedback..");

		int feedbackCounter = 0;
		int debugCounter = 0;

		try {
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));

			// Find and replace in paragraph
			List<XWPFParagraph> paragraphList = document.getParagraphs();

			for (XWPFParagraph paragraph : paragraphList) {
				List<XWPFRun> runs = paragraph.getRuns();

				if (runs != null) {

					for (XWPFRun r : runs) {

						// System.out.println(debugCounter++ + ". "+ r.toString());

						String tempString = r.toString();
						if (isFeedback(tempString)) {
							feedbackCounter++;
							System.out.println(feedbackCounter + ". " + r.toString());
						}
					}
				}
			}

			// close FileInputSTream, XWPFDocument
			fis.close();
			document.close();
		}

		catch (Exception ex) {
			ex.printStackTrace();
		}

		return feedbackCounter;

	}// end of find_Questions

}// end of DocXToXML