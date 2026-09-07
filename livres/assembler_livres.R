# Réassemble les PDF intérieurs des deux livres (à lancer dans RStudio, depuis le dossier "livres").
# 1) Session > Set Working Directory > To Source File Location
# 2) Exécuter tout (Ctrl+Shift+Entrée)
if (!requireNamespace("qpdf", quietly = TRUE)) install.packages("qpdf")
for (livre in c("boulay", "lemonnier")) {
  parties <- sort(list.files(file.path("parts_", livre, fsep = ""), pattern = "\\.pdf$", full.names = TRUE))
  sortie  <- paste0("interieur_", livre, ".pdf")
  qpdf::pdf_combine(parties, output = sortie)
  cat(livre, ":", length(parties), "parties ->", sortie, "(", qpdf::pdf_length(sortie), "pages )\n")
}
