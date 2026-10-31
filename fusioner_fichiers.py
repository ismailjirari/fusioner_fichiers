import os
import subprocess
import tempfile
from pathlib import Path

from PIL import Image
from pypdf import PdfWriter, PdfReader


# ============================================================
# CONFIGURATION
# ============================================================

# Dossier contenant les fichiers
DOSSIER = Path(r"C:\Users\ASUS ZENBOOK\OneDrive\Bureau\projet\fusioner_fichiers")

# Nom du PDF final
PDF_FINAL = DOSSIER / "FUSION_FINAL.pdf"


# Extensions acceptées
PDF_EXTENSIONS = {".pdf"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"}
WORD_EXTENSIONS = {".docx", ".doc"}


# ============================================================
# CONVERSION IMAGE -> PDF
# ============================================================

def image_to_pdf(image_path, output_pdf):
    """
    Convertit une image en PDF.
    L'image est placée sur une page avec sa taille originale.
    """

    image = Image.open(image_path)

    # Gestion des images avec transparence
    if image.mode in ("RGBA", "LA", "P"):
        background = Image.new("RGB", image.size, "white")

        if image.mode == "P":
            image = image.convert("RGBA")

        background.paste(image, mask=image.getchannel("A"))
        image = background
    else:
        image = image.convert("RGB")

    image.save(output_pdf, "PDF", resolution=150.0)

    print(f"[IMAGE -> PDF] {image_path.name}")


# ============================================================
# CONVERSION WORD -> PDF
# ============================================================

def word_to_pdf(word_path, output_folder):
    """
    Convertit un fichier Word en PDF grâce à Microsoft Word.
    """

    output_folder = Path(output_folder)

    # PowerShell lance Microsoft Word via COM
    powershell_script = f'''
$word = New-Object -ComObject Word.Application
$word.Visible = $false

$doc = $word.Documents.Open("{word_path}")

$doc.SaveAs("{output_folder}", 17)

$doc.Close()
$word.Quit()
'''

    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            powershell_script
        ],
        check=True
    )

    pdf_path = output_folder / (word_path.stem + ".pdf")

    print(f"[WORD -> PDF] {word_path.name}")

    return pdf_path


# ============================================================
# PROGRAMME PRINCIPAL
# ============================================================

def main():

    print("=" * 70)
    print("        FUSION PDF + IMAGES + WORD")
    print("=" * 70)

    print(f"\nDossier : {DOSSIER}")

    if not DOSSIER.exists():
        print("\nERREUR : le dossier n'existe pas.")
        return

    # --------------------------------------------------------
    # Création d'un dossier temporaire
    # --------------------------------------------------------

    temp_folder = DOSSIER / "_temp_pdf"

    temp_folder.mkdir(exist_ok=True)

    fichiers_pdf = []

    # --------------------------------------------------------
    # Récupération de tous les fichiers
    # --------------------------------------------------------

    fichiers = sorted(
        [
            f for f in DOSSIER.iterdir()
            if f.is_file()
            and f.name != PDF_FINAL.name
            and not f.name.startswith("~$")
        ],
        key=lambda x: x.name.lower()
    )

    print(f"\nNombre de fichiers trouvés : {len(fichiers)}\n")

    # --------------------------------------------------------
    # Traitement des fichiers
    # --------------------------------------------------------

    for fichier in fichiers:

        extension = fichier.suffix.lower()

        # ================================================
        # PDF
        # ================================================

        if extension in PDF_EXTENSIONS:

            fichiers_pdf.append(fichier)

            print(f"[PDF]        {fichier.name}")

        # ================================================
        # IMAGE
        # ================================================

        elif extension in IMAGE_EXTENSIONS:

            try:

                pdf_image = temp_folder / f"{fichier.stem}_image.pdf"

                image_to_pdf(fichier, pdf_image)

                fichiers_pdf.append(pdf_image)

            except Exception as e:

                print(f"[ERREUR IMAGE] {fichier.name}")
                print(f"              {e}")

        # ================================================
        # WORD
        # ================================================

        elif extension in WORD_EXTENSIONS:

            try:

                pdf_word = temp_folder / f"{fichier.stem}_word.pdf"

                # Conversion Word vers PDF
                word_to_pdf(
                    fichier,
                    pdf_word
                )

                if pdf_word.exists():

                    fichiers_pdf.append(pdf_word)

            except Exception as e:

                print(f"[ERREUR WORD] {fichier.name}")
                print(f"              {e}")

        # ================================================
        # AUTRE
        # ================================================

        else:

            print(f"[IGNORÉ]     {fichier.name}")

    # --------------------------------------------------------
    # Vérification
    # --------------------------------------------------------

    if not fichiers_pdf:

        print("\nAucun fichier PDF, image ou Word trouvé.")
        return

    print("\n" + "=" * 70)
    print("FUSION DES FICHIERS")
    print("=" * 70)

    # --------------------------------------------------------
    # Fusion
    # --------------------------------------------------------

    writer = PdfWriter()

    nombre_pages = 0

    for pdf in fichiers_pdf:

        try:

            reader = PdfReader(str(pdf))

            print(f"\nAjout : {pdf.name}")

            for page in reader.pages:

                writer.add_page(page)
                nombre_pages += 1

        except Exception as e:

            print(f"[ERREUR PDF] {pdf.name}")
            print(f"              {e}")

    # --------------------------------------------------------
    # Création du PDF final
    # --------------------------------------------------------

    with open(PDF_FINAL, "wb") as output:

        writer.write(output)

    # --------------------------------------------------------
    # Nettoyage
    # --------------------------------------------------------

    try:

        for fichier in temp_folder.iterdir():

            fichier.unlink()

        temp_folder.rmdir()

    except Exception:

        pass

    # --------------------------------------------------------
    # Résultat
    # --------------------------------------------------------

    print("\n" + "=" * 70)
    print("              TERMINÉ !")
    print("=" * 70)

    print(f"\nPDF créé :")
    print(PDF_FINAL)

    print(f"\nNombre de fichiers traités : {len(fichiers_pdf)}")
    print(f"Nombre total de pages     : {nombre_pages}")

    print("\nTu peux maintenant ouvrir :")
    print(PDF_FINAL)


# ============================================================
# EXECUTION
# ============================================================

if __name__ == "__main__":
    main()
