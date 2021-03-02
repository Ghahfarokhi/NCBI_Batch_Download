# Retrieve and Annotate Sequences from NCBI using Excel

<a href="https://github.com/Ghahfarokhi/ncbi_batch_download" target="_blank">
<img src="./screenshots/github-repo.png" 
alt="Github Repo"/></a>

A macro-enabled Excel workbook, which could be used to download and annotate sequences for a large set of genomic coordinates or accession numbers from NCBI using Microsoft Excel workbooks. 


<img src="./screenshots/RefSeq_Downloader_Header.jpeg" 
alt="Excel - download and annotate genbank files"/>

---

## Supported systems

* Windows

## Requirements

* Microsoft Excel 2016 or higher

* Internet Connection

## Installation

 * None, just enable macros using the pop up bar, which appears upon opening the workbooks.

## Tools

* **RefSeq-Downloader-v1.xlsm**: downloads genbank files and seuences for a range of genomic coordinates *e.g., Chr1:1000000-1001000*.

* **Accession-Downloader-v1.xlsm**: downloads genbank files and sequences for a list of accession numbers *e.g., NM_000163.5*.

---

## Instruction

#### Youtube

Wath the detailed guideline on youtube:

<a href="http://www.youtube.com/watch?feature=player_embedded&v=T97HnEkpupI" target="_blank">
<img src="./screenshots/youtube.png" 
alt="Youtube instruction" style="width: 500px;"/></a>

#### Step-by-step guideline

1. Download a copy of the excel files on your desktop. Enable the macros using the pop-up menu which appears upon opening the workbooks.

<img src="./screenshots/enable-contents.png" 
alt="enable macros upon opening the excel files" 
style="width: 500px;"/>

2. Populate the user_input table with required genomic coordinates (example input is provided in the ”Info” worksheet). 

* **RefSeq-Downloader-v1.xlsm**: genomic coordinates:

<img src="./screenshots/required-info-genomic-coordinates.jpeg" 
alt="RefSeq-Downloader-v1.xlsm for genomic coordinates" 
style="width: 500px;"/>

* **Accession-Downloader-v1.xlsm**: accession numbers:

<img src="./screenshots/required-info-accession-numbers.jpeg" 
alt="Accession-Downloader-v1.xlsm for accession numbers" 
style="width: 500px;"/>

3. Optionally, a sequence to be annotated on the Genbank file could also be provided:

<img src="./screenshots/optional-inputs.jpeg" 
alt="optional inputs" 
style="width: 500px;"/>

4. Choose if you want (or don’t want) the GenBank files to be saved on your computer and click download. Genbank or fasta files can be found in the same folder where Excel workbooks are stored.

<img src="./screenshots/genbank-file-download-option.png" 
alt="options for downloading" 
style="width: 500px;"/>

5. You can track the progress by checking the Excel status bar. You will be notified when the download is complete. GenBank files will be saved within the folder of the the Excel workbook. Check the "Log" worksheet for the success/failure for each file.

<img src="./screenshots/tracking-the-progress.jpeg" 
alt="tracking the progress" 
style="width: 500px;"/>

6. There are other additional functions available in the workbooks for Reverse Complement and Translations of the sequences:

<img src="./screenshots/other-avaialable-functions.jpeg" 
alt="available functions in Excel workbooks - reverse complement, translate, etc." 
style="width: 500px;"/>

## Limitations
 * The maximum number of genomic coordinates to be downloaded at one go is limited to 1000 to avoid the Excel application running out of memory.

* The maximum length of a genomic region to download is 300,000 bp.

* The maximum length of a sequence in an Excel' cell is 32,767.

* Currently, only a dozen of "popular" assemblies are added (provided in the info worksheet). Watch the instruction on YouTube to learn how to add other genome assemblies from other organisms.

--- 

## Download materials

