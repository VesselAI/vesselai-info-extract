The four Python programs in this project extract relevant pages from PDF files, which contain marine engine data.

Three first Python programs realize a workflow:
VesselAI_ExtractData_1.py
VesselAI_ExtractData_2.py
VesselAI_ExtractData_3.py
In Phase 1 the user enters keywords and other words, and the program finds the matching files/pages; links to these files/pages are saved.
The Phase 2 program reads these files, and tries to group into clusters (4 algorithms); central page in each cluster is saved as a Model Page.
In Phase 3 the user can edit this set of Model Pages; and then the program tries to find more of similar pages, as well as extract relevant lines.

Besides, a separate Neural Network Python program is available:
VesselAI_ExtractData_NN.py
This version uses a part of the input PDF files for teaching, then uses the remaining input PDF files for evaluation. Ground Truth (=relevant pages) is programmed in.

The input PDF files used in program development were as follows:
There are 73 PDF files in total, containing 18 019 pages.
These files are (or at least were in spring 2022) publicly available.
However, due to copyright issues we cannot distribute them.
Lists of these files are presented in the User Guide presentation.
You must try to find them from the vendors’ web sites and download them yourself, as instructed in the User Guide.
The Python programs contain hard-wired lists of the files to be processed. If you cannot find all these listed files, you must modify these lists.
But the NN (neural network) program version also contains the indices of documents and pages that the human expert considered relevant (“Ground Truth”). These index lists must be modified also, which is tedious…

