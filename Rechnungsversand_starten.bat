@echo off
rem Taeglicher Rechnungsversand (Onlineshop + eBay, nie Amazon). Wird von der Windows-Aufgabenplanung um 15:00 gestartet.
rem Siehe RECHNUNGSVERSAND.md. Modus (bericht / entwurf / senden) steht in rechnungsversand_config.json.
cd /d "%~dp0"
py rechnungsversand.py %*
