
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
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import java.io.File;

//https://www.codota.com/code/java/classes/org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun

//https://github.com/Sayi/poi-tl/blob/master/src/main/java/com/deepoove/poi/XWPFParagraphWrapper.java#L49

//https://stackoverflow.com/questions/35088893/apache-poi-remove-cthyperlink-low-level-code/35101440#35101440

//https://www.w3.org/standards/xml/core

//https://docs.moodle.org/38/en/Moodle_XML_format

//https://www.tutorialspoint.com/java_xml/java_dom_create_document.htm

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
	public static boolean isIncorrect(String str) {
		boolean result = false;

		if ( !Character.isDigit(str.charAt(0)) ) {
			if (str.charAt(0) != '*' && str.charAt(0) != '~') {
				result = true;
			}
		}
		return result;
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
	 * Description: Locate feedback in path and return total number of feedbacks
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

	}// end of find_Feedback


	/**
	 * Description: Locate correct answer  in path and return number of correct answers
	 * 
	 * @param path location of file
	 * @return return total text found in link for docx
	 */
	public static int find_Answer(String path) {

		System.out.println("Initalize find_Answer..");

		int answerCounter = 0;
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

						if (isAnswer(tempString)) {
							answerCounter++;
							System.out.println(answerCounter + ". " + r.toString());
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

		return answerCounter;

	}// end of find_Answer


	/**
	 * Description: Locate incorrect answer  in path and return number of incorrect answers
	 * 
	 * @param path location of file
	 * @return return total text found in link for docx
	 */
	public static int find_Incorrect(String path) {

		System.out.println("Initalize find_Incorrect..");

		int incorrectCounter = 0;
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

						if (isIncorrect(tempString)) {
							incorrectCounter++;
							System.out.println(incorrectCounter + ". " + r.toString());
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

		return incorrectCounter;

	}// end of find_Incorrect

	/**
	 * Description: Convert respondus form quiz docx to xml
	 * 
	 * @param path location of file
	 * @return return XML
	 */
	public static void convert_XML(String path) {

		System.out.println("Initalize XML conversion..");

		try {
			DocumentBuilderFactory dbFactory =
			DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.newDocument();
			
			// root element
			Element rootElement = doc.createElement("quiz");
			doc.appendChild(rootElement);
   
			// supercars element
			Element supercar = doc.createElement("supercars");
			rootElement.appendChild(supercar);
   
			// setting attribute to element
			Attr attr = doc.createAttribute("company");
			attr.setValue("Ferrari");
			supercar.setAttributeNode(attr);
   
			// carname element
			Element carname = doc.createElement("carname");
			Attr attrType = doc.createAttribute("type");
			attrType.setValue("formula one");
			carname.setAttributeNode(attrType);
			carname.appendChild(doc.createTextNode("Ferrari 101"));
			supercar.appendChild(carname);
   
			Element carname1 = doc.createElement("carname");
			Attr attrType1 = doc.createAttribute("type");
			attrType1.setValue("sports");
			carname1.setAttributeNode(attrType1);
			carname1.appendChild(doc.createTextNode("Ferrari 202"));
			supercar.appendChild(carname1);
   
			// write the content into xml file
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(new File("C:\\cars.xml"));
			transformer.transform(source, result);
			
			// Output to console for testing
			StreamResult consoleResult = new StreamResult(System.out);
			transformer.transform(source, consoleResult);
		 } catch (Exception e) {
			e.printStackTrace();
		 }

		//int debugCounter = 0;

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

						} else if (isFeedback(tempString)) {

						} else if (isAnswer(tempString)) {

						} else if (isIncorrect(tempString)) {

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

	}// end of convert_XML

	

}// end of DocXToXML