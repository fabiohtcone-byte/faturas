import { GoogleGenAI } from "@google/genai";
async function run() {
  try {
    const ai = new GoogleGenAI({ apiKey: "" });
    await ai.models.generateContent({ model: "gemini-3-flash-preview", contents: "hi" });
  } catch (e: any) {
    console.log("ERROR MESSAGE:", e.message);
  }
}
run();
