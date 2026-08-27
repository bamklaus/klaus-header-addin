async function applyHeaderRule() {
  await Word.run(async (context) => {

    // 1. Læs første afsnit i dokumentet
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.load("text");
    await context.sync();

    const text = firstParagraph.text.trim();

    // 2. Find titel i formatet [Title: X]
    const match = text.match(/

\[Title:\s*(.*?)\]

/i);
    if (!match) return; // Ingen titel fundet

    const title = match[1].trim();

    // 3. Indsæt titel i headeren
    const sections = context.document.sections;
    const header = sections.getFirst().getHeader("Primary");

    header.insertParagraph(title, Word.InsertLocation.replace);

    // 4. Formatering (centreret, Aptos 18)
    const headerParagraph = header.paragraphs.getFirst();
    headerParagraph.font.name = "Aptos";
    headerParagraph.font.size = 18;
    headerParagraph.font.bold = true;
    headerParagraph.alignment = Word.Alignment.center;

    await context.sync();
  });
}
