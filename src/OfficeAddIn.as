package
{
	import org.apache.royale.events.EventDispatcher;

	/**
	 * @externs
	 */
	COMPILE::JS	
	public class OfficeAddIn extends EventDispatcher
	{
		/**
         * <inject_script>
		 * var script = document.createElement("script");
		 * script.setAttribute("src", "resources/office/office.js");
		 * document.head.appendChild(script);
		 * </inject_script>
		 */
		public function OfficeAddIn(){}

		public static function getDefaultLanguage():String { return null; }

		public static function insertText(text:String):void {}

		public static function insertHtml(html:String):void {}

		public static function insertTable(table:Array):void {}

		public static function insertImage(base64Image:String):void {}

		public static function existText(toFind:String):Promise { return null; }

		public static function saveSnapshot():void {}

		public static function restoreSnapshot():void {}

		public static function findAndReplace(toFind:String, toReplace:String):Promise { return null; }

		public static function getDocumentAsPDF():Promise { return null; }
	}
}