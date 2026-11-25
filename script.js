/* ==========================================================================
   DATABASE JADWAL PER STASIUN (FORMAT COPY-PASTE LANGSUNG DARI EXCEL)
   ==========================================================================
   CARA MENAMBAH DATA STASIUN BARU (MENDUKUNG 5 ATAU 6 KOLOM EXCEL):
   - Format 6 Kolom (Dengan Relasi): [No. KA] [NAMA KA] [JAM] [RELASI] [KELAS] [KET]
   - Format 5 Kolom (Tanpa Relasi) : [No. KA] [NAMA KA] [JAM] [KELAS] [KET]
   Paste langsung di dalam tanda backtick (` ... `) pada stasiun yang sesuai.
   Dropdown stasiun, chip akses cepat, dan peta koridor akan otomatis aktif.
   ========================================================================== */

const STATION_META = {
    SOLO_BALAPAN:  { name: "Stasiun Solo Balapan", code: "SLO" },
    YOGYAKARTA:    { name: "Stasiun Yogyakarta", code: "YK" },
    LEMPUYANGAN:   { name: "Stasiun Lempuyangan", code: "LPN" },
    PURWOSARI:     { name: "Stasiun Purwosari", code: "PWS" },
    SOLOJEBRES:    { name: "Stasiun Solo Jebres", code: "SK" },
    KLATEN:        { name: "Stasiun Klaten", code: "KT" },
    WONOGIRI:      { name: "Stasiun Wonogiri", code: "WNG" },
    PALUR:         { name: "Stasiun Palur", code: "PL" },
    GAWOK:         { name: "Stasiun Gawok", code: "GW" },
    DELANGGU:      { name: "Stasiun Delanggu", code: "DL" },
    CEPER:         { name: "Stasiun Ceper", code: "CE" },
    SROWOT:        { name: "Stasiun Srowot", code: "SWT" },
    BRAMBANAN:     { name: "Stasiun Brambanan", code: "BBN" },
    MAGUWO:        { name: "Stasiun Maguwo", code: "MGW" },
    PATUKAN:       { name: "Stasiun Patukan", code: "PTN" },
    REWULU:        { name: "Stasiun Rewulu", code: "RWL" },
    KEDUNDANG:     { name: "Stasiun Kedundang", code: "KDG" },
    WATES:         { name: "Stasiun Wates", code: "WT" },
    KADIPIRO:      { name: "Stasiun Kadipiro", code: "KDO" },
    ADI_SOEMARMO:  { name: "Stasiun Bandara Adi Soemarmo", code: "SMO" },
    SOLOKOTA:      { name: "Stasiun Solo Kota", code: "STA" },
    SUKOHARJO:     { name: "Stasiun Sukoharjo", code: "SKH" },
    PASARNGUTER:   { name: "Stasiun Pasarnguter", code: "PNT" },
    KALIOSO:       { name: "Stasiun Kalioso", code: "KO" },
    SALEM:         { name: "Stasiun Salem", code: "SLM" },
    SUMBERLAWANG:  { name: "Stasiun Sumberlawang", code: "SUM" },
    GOPRAK:        { name: "Stasiun Goprak", code: "GPK" },
    JENAR:         { name: "Stasiun Jenar", code: "JN" },
    WOJO:          { name: "Stasiun Wojo", code: "WJ" },
    SENTOLO:       { name: "Stasiun Sentolo", code: "STL" },
    BANDARA_YIA:   { name: "Stasiun Bandara YIA", code: "YIA" },
    KEMIRI:        { name: "Stasiun Kemiri", code: "KMR" },
    MASARAN:       { name: "Stasiun Masaran", code: "MSR" },
    SRAGEN:        { name: "Stasiun Sragen", code: "SR" },
    KEBONROMO:     { name: "Stasiun Kebonromo", code: "KRO" },
    KEDUNGBANTENG: { name: "Stasiun Kedungbanteng", code: "KDB" }
};

const RAW_STATION_DATABASE = {
    // #DATABASE_STASIUN_SOLO_BALAPAN
    SOLO_BALAPAN: `
8	BIMA	00.07.00	ARGO	BERHENTI
81	SANGKURIANG	00.10.00	EKSEKUTIF EKONOMI	BERHENTI
260F	BANYUBIRU	00.13.00	EKSEKUTIF EKONOMI	TERMINUS
728	COMMUTER LINE YK	00.15.00	KRL	BERHENTI
307	KAHURIPAN	00.21.00	EKONOMI	LANGSUNG
2706F	BRAMNAMBO SERVICE	00.30.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
347	PARCEL SELATAN	00.34.00	KA BARANG PARCEL	LANGSUNG
12	TURANGGA	00.40.00	ARGO	BERHENTI
79F	GAJAYANA FAKULTATIF	00.42.00	EKSEKUTIF	BERHENTI
52	BRAWIJAYA	00.51.00	EKSEKUTIF	BERHENTI
115F	SANCAKA	01.05.00	EKSEKUTIF EKONOMI	BERHENTI
2711	BUNGTALUN SERVICE	01.20.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
51	BRAWIJAYA	01.30.00	EKSEKUTIF	BERHENTI
290	JAYAKARTA	01.33.00	EKONOMI	BERHENTI
86	MALABAR	02.04.00	EKSEKUTIF EKONOMI	BERHENTI
2601	MAWALO TANKER	02.24.00	KA BARANG KHUSUS (BBM)	LANGSUNG
50	GAJAYANA	02.36.00	EKSEKUTIF	BERHENTI
108	LODAYA	03.00.00	EKSEKUTIF EKONOMI	TERMINUS
14	ARGO LAWU	03.42.00	ARGO	TERMINUS
2703F	BRAMNAMBO SERVICE	03.55.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
96	MUTIARA SELATAN	03.56.00	EKSEKUTIF EKONOMI	BERHENTI
2708F	BRAMARJA SERVICE	04.16.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
2709F	BRAMARJA SERVICE	04.25.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
348	PARCEL SELATAN	04.34.00	KA BARANG PARCEL	LANGSUNG
182	MADIUN JAYA	04.42.00	EKSEKUTIF EKONOMI	BERHENTI
701	COMMUTER LINE YK	05.21.00	KRL	BERHENTI
344	PARCEL TENGAH	05.24.00	KA BARANG PARCEL	LANGSUNG
233	JOGLOSEMARKERTO	05.30.00	EKSEKUTIF EKONOMI	START
720	COMMUTER LINE YK	06.03.00	KRL	BERHENTI
186	MATARAM	06.13.00	EKSEKUTIF EKONOMI	TERMINUS
600	BANDARA SMO	06.19.00	KA BANDARA	BERHENTI
239	JOGLOSEMARKERTO	06.20.00	EKSEKUTIF EKONOMI	START
711	COMMUTER LINE YK	06.22.00	KRL	BERHENTI
62	MANAHAN	06.33.00	EKSEKUTIF	TERMINUS
150	SINGASARI	06.51.00	EKSEKUTIF EKONOMI	LANGSUNG
730	COMMUTER LINE YK	07.08.00	KRL	BERHENTI
105	LODAYA	07.10.00	EKSEKUTIF EKONOMI	START
721	COMMUTER LINE YK	07.15.00	KRL	BERHENTI
110	SANCAKA	07.19.00	EKSEKUTIF EKONOMI	BERHENTI
308	KAHURIPAN	07.35.00	EKONOMI	LANGSUNG
280	PANDAWA	07.37.00	EKSEKUTIF	BERHENTI
279	PANDAWA	07.45.00	EKSEKUTIF	BERHENTI
312	SRI TANJUNG	07.50.00	EKONOMI	LANGSUNG
595	BANDARA SMO	07.50.00	KA BANDARA	BERHENTI
731F	COMMUTER LINE YK	07.56.00	KRL	BERHENTI
702	COMMUTER LINE YK	08.02.00	KRL	BERHENTI
13	ARGO LAWU	08.20.00	ARGO	START
80F	GAJAYANA FAKULTATIF	08.25.00	EKSEKUTIF	BERHENTI
338F	BANYUBIRU	08.27.00	EKONOMI	TERMINUS
703	COMMUTER LINE YK	08.36.00	KRL	BERHENTI
596	BANDARA SMO	08.45.00	KA BANDARA	BERHENTI
309	PASUNDAN	08.49.00	EKONOMI	LANGSUNG
712	COMMUTER LINE YK	09.02.00	KRL	BERHENTI
2712	BUNGTALUN SERVICE	09.12.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
337F	BANYUBIRU	09.15.00	EKONOMI	START
181	MADIUN JAYA	09.20.00	EKSEKUTIF EKONOMI	BERHENTI
89	GAYABARU SELATAN	09.32.00	EKSEKUTIF EKONOMI	LANGSUNG
722	COMMUTER LINE YK	09.47.00	KRL	BERHENTI
109	SANCAKA	09.55.00	EKSEKUTIF EKONOMI	BERHENTI
294	LOGAWA	10.01.00	EKSEKUTIF EKONOMI	LANGSUNG
2713	BUNGTALUN SERVICE	10.08.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
59	MANAHAN	10.10.00	EKSEKUTIF	START
87	MALABAR	10.25.00	EKSEKUTIF EKONOMI	BERHENTI
185	MATARAM	10.40.00	EKSEKUTIF EKONOMI	START
732F	COMMUTER LINE YK	10.42.00	KRL	BERHENTI
175	BANGUNKARTA	10.46.00	EKSEKUTIF EKONOMI	BERHENTI
713	COMMUTER LINE YK	10.55.00	KRL	BERHENTI
601	BANDARA SMO	11.11.00	KA BANDARA	BERHENTI
704	COMMUTER LINE YK	11.16.00	KRL	BERHENTI
216F	SANCAKA UTARA	11.17.00	EKSEKUTIF BISNIS	BERHENTI
217F	SANCAKA UTARA	11.22.00	EKSEKUTIF BISNIS	BERHENTI
190	MALIOBORO	11.25.00	EKSEKUTIF EKONOMI	BERHENTI
9	ARGO WILIS	11.33.00	ARGO	BERHENTI
282	PANDAWA	11.35.00	EKSEKUTIF	BERHENTI
281	PANDAWA	11.40.00	EKSEKUTIF	BERHENTI
2704F	BRAMNAMBO SERVICE	12.02.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
112	SANCAKA	12.04.00	EKSEKUTIF EKONOMI	BERHENTI
5	ARGO SEMERU	12.09.00	ARGO	BERHENTI
705	COMMUTER LINE YK	12.25.00	KRL	BERHENTI
170	RANGGAJATI	12.30.00	EKSEKUTIF EKONOMI	BERHENTI
2705F	BRAMNAMBO SERVICE	12.40.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
54	BRAWIJAYA	12.42.00	EKSEKUTIF	BERHENTI
2603	MAWALO TANKER	12.46.00	KA BARANG KHUSUS (BBM)	LANGSUNG
2602	MAWALO TANKER	12.48.00	KA BARANG KHUSUS (BBM)	LANGSUNG
2710F	BRAMARJA SERVICE	13.04.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
723F	COMMUTER LINE YK	13.05.00	KRL	BERHENTI
2707F	BRAMARJA SERVICE	13.15.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
714	COMMUTER LINE YK	13.22.00	KRL	BERHENTI
169	RANGGAJATI	13.26.00	EKSEKUTIF EKONOMI	BERHENTI
6	ARGO SEMERU	13.32.00	ARGO	BERHENTI
602	BANDARA SMO	13.36.00	KA BANDARA	BERHENTI
187	KERTANEGARA	13.47.00	EKSEKUTIF EKONOMI	BERHENTI
184	FAJAR UTAMA SLO	13.50.00	EKSEKUTIF EKONOMI	TERMINUS
715	COMMUTER LINE YK	13.55.00	KRL	BERHENTI
106	LODAYA	14.03.00	EKSEKUTIF EKONOMI	TERMINUS
111	SANCAKA	14.14.00	EKSEKUTIF EKONOMI	BERHENTI
2716	KARSOLO SERVICE	14.15.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
597	BANDARA SMO	14.25.00	KA BANDARA	BERHENTI
10	ARGO WILIS	14.30.00	ARGO	BERHENTI
733F	COMMUTER LINE YK	14.35.00	KRL	BERHENTI
287	MATARMAJA	15.02.00	EKONOMI	LANGSUNG
706	COMMUTER LINE YK	15.06.00	KRL	BERHENTI
598	BANDARA SMO	15.19.00	KA BANDARA	BERHENTI
16	ARGO DWIPANGGA	15.27.00	ARGO	TERMINUS
707	COMMUTER LINE YK	15.38.00	KRL	BERHENTI
724F	COMMUTER LINE YK	15.58.00	KRL	BERHENTI
53	BRAWIJAYA	16.08.00	EKSEKUTIF	BERHENTI
234	JOGLOSEMARKERTO	16.11.00	EKSEKUTIF EKONOMI	TERMINUS
340	PANDAWA	16.28.00	EKSEKUTIF	BERHENTI
235	JOGLOSEMARKERTO	16.34.00	EKSEKUTIF EKONOMI	TERMINUS
339	PANDAWA	16.35.00	EKSEKUTIF	BERHENTI
725	COMMUTER LINE YK	16.38.00	KRL	BERHENTI
716	COMMUTER LINE YK	16.44.00	KRL	BERHENTI
298	JAKA TINGKIR	17.05.00	EKONOMI	TERMINUS
293	LOGAWA	17.10.00	EKSEKUTIF EKONOMI	LANGSUNG
88	MALABAR	17.13.00	EKSEKUTIF EKONOMI	BERHENTI
238	JOGLOSEMARKERTO	17.20.00	EKSEKUTIF EKONOMI	TERMINUS
717	COMMUTER LINE YK	17.20.00	KRL	BERHENTI
734F	COMMUTER LINE YK	17.32.00	KRL	BERHENTI
151	BRANTAS	17.42.00	EKSEKUTIF EKONOMI	LANGSUNG
114	SANCAKA	17.46.00	EKSEKUTIF EKONOMI	BERHENTI
126	BATAVIA	18.01.00	EKSEKUTIF EKONOMI	BERHENTI
183	SENJA UTAMA SLO	18.05.00	EKSEKUTIF EKONOMI	START
289	JAYAKARTA	18.09.00	EKONOMI	BERHENTI
259F	BANYUBIRU	18.15.00	EKSEKUTIF EKONOMI	START
603	BANDARA SMO	18.20.00	KA BANDARA	BERHENTI
2715	KARSOLO SERVICE	18.20.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
60	MANAHAN	18.26.00	EKSEKUTIF	TERMINUS
311	SRI TANJUNG	18.32.00	EKONOMI	LANGSUNG
708	COMMUTER LINE YK	18.34.00	KRL	BERHENTI
735	COMMUTER LINE YK	18.46.00	KRL	BERHENTI
726	COMMUTER LINE YK	19.09.00	KRL	BERHENTI
604	BANDARA SMO	19.12.00	KA BANDARA	BERHENTI
172	WIJAYA KUSUMA	19.19.00	EKSEKUTIF EKONOMI	BERHENTI
107	LODAYA	19.20.00	EKSEKUTIF EKONOMI	START
342	PANDAWA	19.31.00	EKSEKUTIF	BERHENTI
113	SANCAKA	19.31.00	EKSEKUTIF EKONOMI	BERHENTI
341	PANDAWA	19.35.00	EKSEKUTIF	BERHENTI
599	BANDARA SMO	19.39.00	KA BANDARA	BERHENTI
49	GAJAYANA	19.53.00	EKSEKUTIF	BERHENTI
90	GAYABARU SELATAN	19.53.00	EKSEKUTIF EKONOMI	LANGSUNG
709	COMMUTER LINE YK	20.00.00	KRL	BERHENTI
310	PASUNDAN	20.07.00	EKONOMI	LANGSUNG
718	COMMUTER LINE YK	20.17.00	KRL	BERHENTI
15	ARGO DWIPANGGA	20.25.00	ARGO	START
288	MATARMAJA	20.34.00	EKONOMI	LANGSUNG
149	SINGASARI	20.46.00	EKSEKUTIF EKONOMI	LANGSUNG
171	WIJAYA KUSUMA	20.57.00	EKSEKUTIF EKONOMI	BERHENTI
218F	SANCAKA UTARA	21.01.00	EKSEKUTIF BISNIS	BERHENTI
219F	SANCAKA UTARA	21.08.00	EKSEKUTIF BISNIS	BERHENTI
95	MUTIARA SELATAN	21.11.00	EKSEKUTIF EKONOMI	BERHENTI
176	BANGUNKARTA	21.13.00	EKSEKUTIF EKONOMI	BERHENTI
727	COMMUTER LINE YK	21.20.00	KRL	BERHENTI
736	COMMUTER LINE YK	21.43.00	KRL	BERHENTI
85	MALABAR	21.44.00	EKSEKUTIF EKONOMI	BERHENTI
286	MAJAPAHIT	21.53.00	EKONOMI	LANGSUNG
125	BATAVIA	22.05.00	EKSEKUTIF EKONOMI	START
7	BIMA	22.13.00	ARGO	BERHENTI
188	KERTANEGARA	22.14.00	EKSEKUTIF EKONOMI	BERHENTI
343	PARCEL TENGAH	22.24.00	KA BARANG PARCEL	LANGSUNG
61	MANAHAN	22.35.00	EKSEKUTIF	TERMINUS
710F	COMMUTER LINE YK	22.44.00	KRL	BERHENTI
2714	BUNGTALUN SERVICE	22.51.00	KA BARANG KHUSUS (SEMEN)	TERMINUS
11	TURANGGA	23.01.00	ARGO	BERHENTI
116F	SANCAKA	23.01.00	EKSEKUTIF EKONOMI	BERHENTI
285	MAJAPAHIT	23.08.00	EKONOMI	LANGSUNG
297	JAKA TINGKIR	23.10.00	EKONOMI	START
2604	MAWALO TANKER	23.19.00	KA BARANG KHUSUS (BBM)	LANGSUNG
232	JOGLOSEMARKERTO	23.27.00	EKSEKUTIF EKONOMI	TERMINUS
82	SANGKURIANG	23.31.00	EKSEKUTIF EKONOMI	BERHENTI
189	MALIOBORO	23.47.00	EKSEKUTIF EKONOMI	BERHENTI
152	BRANTAS	23.50.00	EKSEKUTIF EKONOMI	LANGSUNG
`,

    // #DATABASE_STASIUN_YOGYAKARTA
    YOGYAKARTA: `
5	ARGO SEMERU	12.52.00	Surabayagubeng - Gambir	ARGO	BERHENTI
6	ARGO SEMERU	12.46.00	Gambir - Surabayagubeng	ARGO	BERHENTI
7	BIMA	22.56.00	Surabayagubeng - Gambir	ARGO	BERHENTI
8	BIMA	23.22.00	Gambir - Surabayagubeng	ARGO	BERHENTI
9	ARGO WILIS	12.17.00	Surabayagubeng - Bandung	ARGO	BERHENTI
10	ARGO WILIS	13.47.00	Bandung - Surabayagubeng	ARGO	BERHENTI
11	TURANGGA	23.44.00	Surabayagubeng - Bandung	ARGO	BERHENTI
12	TURANGGA	23.56.00	Bandung - Surabayagubeng	ARGO	BERHENTI
13	ARGO LAWU	09.05.00	Solobalapan - Gambir	ARGO	BERHENTI
14	ARGO LAWU	02.48.00	Gambir - Solobalapan	ARGO	BERHENTI
15	ARGO DWIPANGGA	21.10.00	Solobalapan - Gambir	ARGO	BERHENTI
16	ARGO DWIPANGGA	14.35.00	Gambir - Solobalapan	ARGO	BERHENTI
29	TAKSAKA	08.45.00	Yogyakarta - Gambir	ARGO	START
30	TAKSAKA	19.21.00	Gambir - Yogyakarta	ARGO	TERMINUS
31	TAKSAKA	21.40.00	Yogyakarta - Gambir	ARGO	START
32	TAKSAKA	13.16.00	Gambir - Yogyakarta	ARGO	TERMINUS
33	TAKSAKA	15.40.00	Yogyakarta - Gambir	ARGO	START
34	TAKSAKA	04.22.00	Gambir - Yogyakarta	ARGO	TERMINUS
49	GAJAYANA	20.37.00	Malang - Gambir	EKSEKUTIF	BERHENTI
50	GAJAYANA	01.48.00	Gambir - Malang	EKSEKUTIF	BERHENTI
59	MANAHAN	10.55.00	Solobalapan - Gambir	EKSEKUTIF	BERHENTI
60	MANAHAN	17.31.00	Gambir - Solobalapan	EKSEKUTIF	BERHENTI
61	MANAHAN	23.21.00	Solobalapan - Gambir	EKSEKUTIF	BERHENTI
62	MANAHAN	05.40.00	Gambir - Solobalapan	EKSEKUTIF	BERHENTI
75F	BOROBUDUR	06.15.00	Lempuyangan - Gambir	EKSEKUTIF	BERHENTI
76F	BOROBUDUR	00.26.00	Gambir - Lempuyangan	EKSEKUTIF	BERHENTI
77F	BOROBUDUR	16.25.00	Lempuyangan - Gambir	EKSEKUTIF	BERHENTI
78F	BOROBUDUR	12.32.00	Gambir - Lempuyangan	EKSEKUTIF	BERHENTI
79F	GAJAYANA FAKULTATIF	01.34.00	Malang - Gambir	EKSEKUTIF	BERHENTI
80F	GAJAYANA FAKULTATIF	07.28.00	Gambir - Malang	EKSEKUTIF	BERHENTI
279	PANDAWA	08.34.00	Semarang Tawang - Solobalapan - Yogyakarta	EKSEKUTIF	TERMINUS
282	PANDAWA	10.50.00	Yogyakarta - Solobalapan - Semarang Tawang	EKSEKUTIF	START
339	PANDAWA	17.22.00	Semarang Tawang - Solobalapan - Yogyakarta	EKSEKUTIF	TERMINUS
342	PANDAWA	18.35.00	Yogyakarta - Solobalapan - Semarang Tawang	EKSEKUTIF	START
82	SANGKURIANG	22.45.00	Bandung - Surabayagubeng - Ketapang	EKSEKUTIF EKONOMI	BERHENTI
81	SANGKURIANG	00.56.00	Ketapang - Surabayagubeng - Bandung	EKSEKUTIF EKONOMI	BERHENTI
85	MALABAR	22.31.00	Malang - Bandung	EKSEKUTIF EKONOMI	BERHENTI
86	MALABAR	01.10.00	Bandung - Malang	EKSEKUTIF EKONOMI	BERHENTI
87	MALABAR	11.14.00	Malang - Bandung	EKSEKUTIF EKONOMI	BERHENTI
88	MALABAR	16.21.00	Bandung - Malang	EKSEKUTIF EKONOMI	BERHENTI
89	GAYABARU SELATAN	10.27.00	Surabayagubeng - Pasarsenen	EKSEKUTIF EKONOMI	LANGSUNG
90	GAYABARU SELATAN	18.52.00	Pasarsenen - Surabayagubeng	EKSEKUTIF EKONOMI	LANGSUNG
95	MUTIARA SELATAN	21.58.00	Surabayagubeng - Bandung	EKSEKUTIF EKONOMI	BERHENTI
96	MUTIARA SELATAN	03.10.00	Bandung - Surabayagubeng	EKSEKUTIF EKONOMI	BERHENTI
105	LODAYA	07.56.00	Solobalapan - Bandung	EKSEKUTIF EKONOMI	BERHENTI
106	LODAYA	13.07.00	Bandung - Solobalapan	EKSEKUTIF EKONOMI	BERHENTI
107	LODAYA	20.06.00	Solobalapan - Bandung	EKSEKUTIF EKONOMI	BERHENTI
108	LODAYA	02.04.00	Bandung - Solobalapan	EKSEKUTIF EKONOMI	BERHENTI
109	SANCAKA	10.47.00	Surabayagubeng - Yogyakarta	EKSEKUTIF EKONOMI	TERMINUS
110	SANCAKA	06.35.00	Yogyakarta - Surabayagubeng	EKSEKUTIF EKONOMI	START
111	SANCAKA	15.07.00	Surabayagubeng - Yogyakarta	EKSEKUTIF EKONOMI	TERMINUS
112	SANCAKA	11.20.00	Yogyakarta - Surabayagubeng	EKSEKUTIF EKONOMI	START
113	SANCAKA	20.23.00	Surabayagubeng - Yogyakarta	EKSEKUTIF EKONOMI	TERMINUS
114	SANCAKA	17.00.00	Yogyakarta - Surabayagubeng	EKSEKUTIF EKONOMI	START
115F	SANCAKA	01.59.00	Surabayagubeng - Yogyakarta	EKSEKUTIF EKONOMI	TERMINUS
116F	SANCAKA	22.15.00	Yogyakarta - Surabayagubeng	EKSEKUTIF EKONOMI	START
125	BATAVIA	23.05.00	Solobalapan - Gambir	EKSEKUTIF EKONOMI	BERHENTI
126	BATAVIA	17.04.00	Gambir - Solobalapan	EKSEKUTIF EKONOMI	BERHENTI
149	SINGASARI	21.49.00	Blitar - Pasarsenen	EKSEKUTIF EKONOMI	LANGSUNG
150	SINGASARI	05.31.00	Pasarsenen - Blitar	EKSEKUTIF EKONOMI	LANGSUNG
169	RANGGAJATI	14.19.00	Jember - Surabayagubeng - Cirebon	EKSEKUTIF EKONOMI	BERHENTI
170	RANGGAJATI	11.38.00	Cirebon - Surabayagubeng - Jember	EKSEKUTIF EKONOMI	BERHENTI
172	WIJAYA KUSUMA	18.25.00	Cilacap - Surabayagubeng - Ketapang	EKSEKUTIF EKONOMI	BERHENTI
171	WIJAYA KUSUMA	22.07.00	Ketapang - Surabayagubeng - Cilacap	EKSEKUTIF EKONOMI	BERHENTI
175	BANGUNKARTA	11.41.00	Jombang - Pasarsenen	EKSEKUTIF EKONOMI	BERHENTI
176	BANGUNKARTA	20.19.00	Pasarsenen - Jombang	EKSEKUTIF EKONOMI	BERHENTI
181	MADIUN JAYA	10.05.00	Madiun - Pasarsenen	EKSEKUTIF EKONOMI	BERHENTI
182	MADIUN JAYA	03.57.00	Pasarsenen - Jombang	EKSEKUTIF EKONOMI	BERHENTI
183	SENJA UTAMA SLO	18.54.00	Solobalapan - Pasarsenen	EKSEKUTIF EKONOMI	BERHENTI
184	FAJAR UTAMA SLO	12.56.00	Pasarsenen - Solobalapan	EKSEKUTIF EKONOMI	BERHENTI
185	MATARAM	11.29.00	Solobalapan - Pasarsenen	EKSEKUTIF EKONOMI	BERHENTI
186	MATARAM	05.16.00	Pasarsenen - Solobalapan	EKSEKUTIF EKONOMI	BERHENTI
187	KERTANEGARA	14.42.00	Malang - Purwokerto	EKSEKUTIF EKONOMI	BERHENTI
188	KERTANEGARA	21.19.00	Purwokerto - Malang	EKSEKUTIF EKONOMI	BERHENTI
189	MALIOBORO	00.41.00	Malang - Purwokerto	EKSEKUTIF EKONOMI	BERHENTI
190	MALIOBORO	10.30.00	Purwokerto - Malang	EKSEKUTIF EKONOMI	BERHENTI
191	BOGOWONTO	08.55.00	Lempuyangan - Pasarsenen	EKSEKUTIF EKONOMI	BERHENTI
192	BOGOWONTO	01.32.00	Pasarsenen - Lempuyangan	EKSEKUTIF EKONOMI	BERHENTI
193	GAJAHWONG	20.54.00	Lempuyangan - Pasarsenen	EKSEKUTIF EKONOMI	LANGSUNG
194	GAJAHWONG	15.45.00	Pasarsenen - Lempuyangan	EKSEKUTIF EKONOMI	LANGSUNG
195	SENJA UTAMA YK	17.50.00	Yogyakarta - Pasarsenen	EKSEKUTIF EKONOMI	START
196	SENJA UTAMA YK	02.27.00	Pasarsenen - Yogyakarta	EKSEKUTIF EKONOMI	TERMINUS
197	FAJAR UTAMA YK	07.00.00	Yogyakarta - Pasarsenen	EKSEKUTIF EKONOMI	START
198	FAJAR UTAMA YK	15.13.00	Pasarsenen - Yogyakarta	EKSEKUTIF EKONOMI	TERMINUS
217F	SANCAKA UTARA	12.29.00	Surabaya Pasarturi - Gambringan - Solobalapan - Cilacap	EKSEKUTIF BISNIS	BERHENTI
218F	SANCAKA UTARA	20.06.00	Cilacap - Solobalapan - Gambringan - Surabaya Pasarturi	EKSEKUTIF BISNIS	BERHENTI
234	JOGLOSEMARKERTO	15.03.00	Solobalapan - Tegal - Solobalapan - Semarang Tawang	EKSEKUTIF EKONOMI	BERHENTI
239	JOGLOSEMARKERTO	07.16.00	Solobalapan - Tegal - Solobalapan	EKSEKUTIF EKONOMI	BERHENTI
247	JOGLOSEMARKERTO	10.45.00	Yogyakarta - Cilacap	EKSEKUTIF EKONOMI	TERMINUS
248	JOGLOSEMARKERTO	09.10.00	Cilacap - Yogyakarta	EKSEKUTIF EKONOMI	TERMINUS
293	LOGAWA	18.14.00	Ketapang - Surabayagubeng - Purwokerto	EKSEKUTIF EKONOMI	LANGSUNG
294	LOGAWA	09.00.00	Purwokerto - Surabayagubeng - Ketapang	EKSEKUTIF EKONOMI	LANGSUNG
289	JAYAKARTA	19.10.00	Surabayagubeng - Pasarsenen	EKSEKUTIF EKONOMI	LANGSUNG
290	JAYAKARTA	00.41.00	Pasarsenen - Surabayagubeng	EKONOMI	LANGSUNG
297	JAKA TINGKIR	00.09.00	Solobalapan - Pasarsenen	EKONOMI	LANGSUNG
298	JAKA TINGKIR	16.04.00	Pasarsenen - Solobalapan	EKONOMI	LANGSUNG
299	PROGO	13.34.00	Lempuyangan - Pasarsenen	EKONOMI	LANGSUNG
300	PROGO	07.19.00	Pasarsenen - Lempuyangan	EKONOMI	LANGSUNG
307	KAHURIPAN	01.23.00	Blitar - Kiaracondong	EKONOMI	LANGSUNG
308	KAHURIPAN	06.13.00	Kiaracondong - Blitar	EKONOMI	LANGSUNG
309	PASUNDAN	09.52.00	Surabayagubeng - Kiaracondong	EKONOMI	LANGSUNG
310	PASUNDAN	18.42.00	Kiaracondong - Surabayagubeng	EKONOMI	LANGSUNG
315	BENGAWAN	00.30.00	Purwosari - Pasarsenen	EKONOMI	LANGSUNG
316	BENGAWAN	17.45.00	Pasarsenen - Purwosari	EKONOMI	LANGSUNG
701	COMMUTER LINE YK	06.33.00	Palur - Yogyakarta	KRL	TERMINUS
702	COMMUTER LINE YK	06.55.00	Yogyakarta - Palur	KRL	START
703	COMMUTER LINE YK	09.57.00	Palur - Yogyakarta	KRL	TERMINUS
704	COMMUTER LINE YK	10.07.00	Yogyakarta - Palur	KRL	START
705	COMMUTER LINE YK	13.42.00	Palur - Yogyakarta	KRL	TERMINUS
706	COMMUTER LINE YK	13.57.00	Yogyakarta - Palur	KRL	START
707	COMMUTER LINE YK	16.49.00	Palur - Yogyakarta	KRL	TERMINUS
708	COMMUTER LINE YK	17.15.00	Yogyakarta - Palur	KRL	START
709	COMMUTER LINE YK	21.23.00	Palur - Yogyakarta	KRL	TERMINUS
710F	COMMUTER LINE YK	21.33.00	Yogyakarta - Palur	KRL	START
711	COMMUTER LINE YK	07.33.00	Palur - Yogyakarta	KRL	TERMINUS
712	COMMUTER LINE YK	07.54.00	Yogyakarta - Palur	KRL	START
713	COMMUTER LINE YK	12.04.00	Palur - Yogyakarta	KRL	TERMINUS
714	COMMUTER LINE YK	12.14.00	Yogyakarta - Palur	KRL	START
715	COMMUTER LINE YK	15.22.00	Palur - Yogyakarta	KRL	TERMINUS
716	COMMUTER LINE YK	15.35.00	Yogyakarta - Palur	KRL	START
717	COMMUTER LINE YK	18.33.00	Palur - Yogyakarta	KRL	TERMINUS
718	COMMUTER LINE YK	19.08.00	Yogyakarta - Palur	KRL	START
720	COMMUTER LINE YK	04.55.00	Yogyakarta - Palur	KRL	START
721	COMMUTER LINE YK	08.26.00	Palur - Yogyakarta	KRL	TERMINUS
722	COMMUTER LINE YK	08.39.00	Yogyakarta - Palur	KRL	START
723F	COMMUTER LINE YK	14.30.00	Palur - Yogyakarta	KRL	TERMINUS
724F	COMMUTER LINE YK	14.50.00	Yogyakarta - Palur	KRL	START
725	COMMUTER LINE YK	17.49.00	Palur - Yogyakarta	KRL	TERMINUS
726	COMMUTER LINE YK	17.58.00	Yogyakarta - Palur	KRL	START
727	COMMUTER LINE YK	22.40.00	Palur - Yogyakarta	KRL	TERMINUS
728	COMMUTER LINE YK	22.55.00	Yogyakarta - Palur	KRL	START
730	COMMUTER LINE YK	06.00.00	Yogyakarta - Palur	KRL	START
731F	COMMUTER LINE YK	09.24.00	Palur - Yogyakarta	KRL	TERMINUS
732F	COMMUTER LINE YK	09.34.00	Yogyakarta - Palur	KRL	START
733F	COMMUTER LINE YK	15.49.00	Palur - Yogyakarta	KRL	TERMINUS
734F	COMMUTER LINE YK	16.14.00	Yogyakarta - Palur	KRL	START
735	COMMUTER LINE YK	19.58.00	Palur - Yogyakarta	KRL	TERMINUS
736	COMMUTER LINE YK	20.33.00	Yogyakarta - Palur	KRL	START
343	PARCEL TENGAH	00.17.00	Malang - Kampungbandan	KA BARANG PARCEL	LANGSUNG
344	PARCEL TENGAH	03.42.00	Kampungbandan - Malang	KA BARANG PARCEL	LANGSUNG
347	PARCEL SELATAN	02.19.00	Surabayakota - Bandung	KA BARANG PARCEL	LANGSUNG
348	PARCEL SELATAN	03.01.00	Bandung - Surabayakota	KA BARANG PARCEL	LANGSUNG
2601	MAWALO TANKER	03.54.00	Madiun - Rewulu	KA BARANG KHUSUS (BBM)	LANGSUNG
2602	MAWALO TANKER	11.02.00	Rewulu - Madiun	KA BARANG KHUSUS (BBM)	LANGSUNG
2603	MAWALO TANKER	14.11.00	Madiun - Rewulu	KA BARANG KHUSUS (BBM)	LANGSUNG
2604	MAWALO TANKER	21.47.00	Rewulu - Madiun	KA BARANG KHUSUS (BBM)	LANGSUNG
515	COMMUTER LINE PRAMEKS	06.40.00	Yogyakarta - Kutoarjo	KRD EKONOMI	START
516	COMMUTER LINE PRAMEKS	09.30.00	Kutoarjo - Yogyakarta	KRD EKONOMI	TERMINUS
517	COMMUTER LINE PRAMEKS	10.16.00	Yogyakarta - Kutoarjo	KRD EKONOMI	START
518	COMMUTER LINE PRAMEKS	14.15.00	Kutoarjo - Yogyakarta	KRD EKONOMI	TERMINUS
519	COMMUTER LINE PRAMEKS	15.00.00	Yogyakarta - Kutoarjo	KRD EKONOMI	START
520	COMMUTER LINE PRAMEKS	18.05.00	Kutoarjo - Yogyakarta	KRD EKONOMI	TERMINUS
521	COMMUTER LINE PRAMEKS	19.23.00	Yogyakarta - Kebumen	KRD EKONOMI	START
522	COMMUTER LINE PRAMEKS	06.28.00	Kebumen - Yogyakarta	KRD EKONOMI	TERMINUS
523	COMMUTER LINE PRAMEKS	16.55.00	Yogyakarta - Kutoarjo	KRD EKONOMI	START
524	COMMUTER LINE PRAMEKS	19.56.00	Kutoarjo - Yogyakarta	KRD EKONOMI	TERMINUS
541	COMMUTER LINE YIA	04.28.00	Yogyakarta - YIA	KA BANDARA	START
542	COMMUTER LINE YIA	05.58.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
543	COMMUTER LINE YIA	06.29.00	Yogyakarta - YIA	KA BANDARA	START
544	COMMUTER LINE YIA	08.00.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
545	COMMUTER LINE YIA	08.19.00	Yogyakarta - YIA	KA BANDARA	START
546	COMMUTER LINE YIA	10.20.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
547	COMMUTER LINE YIA	10.35.00	Yogyakarta - YIA	KA BANDARA	START
548	COMMUTER LINE YIA	12.22.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
549	COMMUTER LINE YIA	14.38.00	Yogyakarta - YIA	KA BANDARA	START
550	COMMUTER LINE YIA	16.37.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
551	COMMUTER LINE YIA	16.45.00	Yogyakarta - YIA	KA BANDARA	START
552	COMMUTER LINE YIA	18.16.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
553	COMMUTER LINE YIA	19.36.00	Yogyakarta - YIA	KA BANDARA	START
554	COMMUTER LINE YIA	21.08.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
555	COMMUTER LINE YIA	07.35.00	Yogyakarta - YIA	KA BANDARA	START
556	COMMUTER LINE YIA	09.19.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
557	COMMUTER LINE YIA	09.40.00	Yogyakarta - YIA	KA BANDARA	START
558	COMMUTER LINE YIA	11.12.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
559	COMMUTER LINE YIA	12.47.00	Yogyakarta - YIA	KA BANDARA	START
560	COMMUTER LINE YIA	14.24.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
561	COMMUTER LINE YIA	15.26.00	Yogyakarta - YIA	KA BANDARA	START
562	COMMUTER LINE YIA	16.53.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
563	COMMUTER LINE YIA	17.10.00	Yogyakarta - YIA	KA BANDARA	START
564	COMMUTER LINE YIA	19.03.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
565	COMMUTER LINE YIA	20.00.00	Yogyakarta - YIA	KA BANDARA	START
566	COMMUTER LINE YIA	21.30.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
567	COMMUTER LINE YIA	06.00.00	Yogyakarta - YIA	KA BANDARA	START
568	COMMUTER LINE YIA	07.40.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
569	COMMUTER LINE YIA	07.50.00	Yogyakarta - YIA	KA BANDARA	START
570	COMMUTER LINE YIA	09.45.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
571	COMMUTER LINE YIA	11.57.00	Yogyakarta - YIA	KA BANDARA	START
572	COMMUTER LINE YIA	13.38.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
573	COMMUTER LINE YIA	13.55.00	Yogyakarta - YIA	KA BANDARA	START
574	COMMUTER LINE YIA	15.30.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
575	COMMUTER LINE YIA	15.55.00	Yogyakarta - YIA	KA BANDARA	START
576	COMMUTER LINE YIA	17.19.00	Yogyakarta - YIA	KA BANDARA	START
577	COMMUTER LINE YIA	17.34.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
578	COMMUTER LINE YIA	19.43.00	Yogyakarta - YIA	KA BANDARA	START
579	COMMUTER LINE YIA	20.25.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
580	COMMUTER LINE YIA	22.07.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
581	COMMUTER LINE YIA	05.10.00	Yogyakarta - YIA	KA BANDARA	START
582	COMMUTER LINE YIA	06.45.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
583	COMMUTER LINE YIA	09.20.00	Yogyakarta - YIA	KA BANDARA	START
584	COMMUTER LINE YIA	11.54.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
585	COMMUTER LINE YIA	13.20.00	Yogyakarta - YIA	KA BANDARA	START
586	COMMUTER LINE YIA	14.54.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
587	COMMUTER LINE YIA	15.08.00	Yogyakarta - YIA	KA BANDARA	START
588	COMMUTER LINE YIA	17.56.00	Yogyakarta - YIA	KA BANDARA	START
589	COMMUTER LINE YIA	18.44.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
590	COMMUTER LINE YIA	20.37.00	YIA - Yogyakarta	KA BANDARA	TERMINUS
2712	BUNGTALUN SERVICE	07.03.00	Karangtalun - Solobalapan - Brumbung	KA BARANG KHUSUS (SEMEN)	LANGSUNG
2711	BUNGTALUN SERVICE	03.14.00	Brumbung - Solobalapan - Karangtalun	KA BARANG KHUSUS (SEMEN)	LANGSUNG
2715	KARSOLO SERVICE	19.47.00	Solobalapan - Karangtalun	KA BARANG KHUSUS (SEMEN)	LANGSUNG
2716	KARSOLO SERVICE	12.05.00	Karangtalun - Solobalapan	KA BARANG KHUSUS (SEMEN)	LANGSUNG
2721	KALEM SERVICE	03.45.00	Lempuyangan - Karangtalun	KA BARANG KHUSUS (SEMEN)	LANGSUNG
2722	KALEM SERVICE	19.34.00	Karangtalun - Lempuyangan	KA BARANG KHUSUS (SEMEN)	LANGSUNG
2750F	CILACE TRANS	04.41.00	Cilacap - Ceper	KA BARANG KHUSUS (PUPUK)	LANGSUNG
2751F	CILACE TRANS	17.59.00	Ceper - Cilacap	KA BARANG KHUSUS (PUPUK)	LANGSUNG
`,

    // #DATABASE_STASIUN_LEMPUYANGAN
    LEMPUYANGAN: ``,
    // #DATABASE_STASIUN_PURWOSARI
    PURWOSARI: ``,
    // #DATABASE_STASIUN_SOLOJEBRES
    SOLOJEBRES: ``,
    // #DATABASE_STASIUN_KLATEN
    KLATEN: ``,
    // #DATABASE_STASIUN_WONOGIRI
    WONOGIRI: ``,
    // #DATABASE_STASIUN_PALUR
    PALUR: ``,
    // #DATABASE_STASIUN_GAWOK
    GAWOK: ``,
    // #DATABASE_STASIUN_DELANGGU
    DELANGGU: ``,
    // #DATABASE_STASIUN_CEPER
    CEPER: ``,
    // #DATABASE_STASIUN_SROWOT
    SROWOT: ``,
    // #DATABASE_STASIUN_BRAMBANAN
    BRAMBANAN: ``,
    // #DATABASE_STASIUN_MAGUWO
    MAGUWO: ``,
    // #DATABASE_STASIUN_PATUKAN
    PATUKAN: ``,
    // #DATABASE_STASIUN_REWULU
    REWULU: ``,
    // #DATABASE_STASIUN_KEDUNDANG
    KEDUNDANG: ``,
    // #DATABASE_STASIUN_WATES
    WATES: ``,
    // #DATABASE_STASIUN_KADIPIRO
    KADIPIRO: ``,
    // #DATABASE_STASIUN_ADI_SOEMARMO
    ADI_SOEMARMO: ``,
    // #DATABASE_STASIUN_SOLOKOTA
    SOLOKOTA: ``,
    // #DATABASE_STASIUN_SUKOHARJO
    SUKOHARJO: ``,
    // #DATABASE_STASIUN_PASARNGUTER
    PASARNGUTER: ``,
    // #DATABASE_STASIUN_KALIOSO
    KALIOSO: `
    51	BRAWIJAYA	01.43.00	Malang - Gambir	EKSEKUTIF	LANGSUNG
    52	BRAWIJAYA	00.38.00	Gambir - Malang	EKSEKUTIF	LANGSUNG
    53	BRAWIJAYA	16.21.00	Malang - Gambir	EKSEKUTIF	LANGSUNG
    54	BRAWIJAYA	12.29.00	Gambir - Malang	EKSEKUTIF	LANGSUNG
    280	PANDAWA	07.23.00	Semarang Tawang - Solobalapan - Yogyakarta	EKSEKUTIF	LANGSUNG
    281	PANDAWA	11.53.00	Yogyakarta - Solobalapan - Semarang Tawang	EKSEKUTIF	LANGSUNG
    340	PANDAWA	16.14.00	Semarang Tawang - Solobalapan - Yogyakarta	EKSEKUTIF	LANGSUNG
    341	PANDAWA	19.48.00	Yogyakarta - Solobalapan - Semarang Tawang	EKSEKUTIF	LANGSUNG
    151	BRANTAS	17.54.00	Blitar - Pasarsenen	EKSEKUTIF EKONOMI	LANGSUNG
    152	BRANTAS	23.37.00	Pasarsenen - Blitar	EKSEKUTIF EKONOMI	LANGSUNG
    216F	SANCAKA UTARA	11.04.00	Surabaya Pasarturi - Gambringan - Solobalapan - Cilacap	EKSEKUTIF BISNIS	LANGSUNG
    219F	SANCAKA UTARA	21.21.00	Cilacap - Solobalapan - Gambringan - Surabaya Pasarturi	EKSEKUTIF BISNIS	LANGSUNG
    232	JOGLOSEMARKERTO	23.12.00	Purwokerto - Tegal - Solobalapan	EKSEKUTIF EKONOMI	LANGSUNG
    233	JOGLOSEMARKERTO	05.43.00	Solobalapan - Tegal - Solobalapan - Semarang Tawang	EKSEKUTIF EKONOMI	LANGSUNG
    235	JOGLOSEMARKERTO	16.47.00	Solobalapan - Tegal - Solobalapan - Semarang Tawang	EKSEKUTIF EKONOMI	LANGSUNG
    238	JOGLOSEMARKERTO	17.06.00	Solobalapan - Tegal - Solobalapan	EKSEKUTIF EKONOMI	LANGSUNG
    259F	BANYUBIRU	18.29.00	Solobalapan - Semarang Tawang	EKSEKUTIF EKONOMI	LANGSUNG
    260F	BANYUBIRU	23.59.00	Semarang Tawang - Solobalapan	EKSEKUTIF EKONOMI	LANGSUNG
    285	MAJAPAHIT	23.20.00	Malang - Pasarsenen	EKONOMI	LANGSUNG
    286	MAJAPAHIT	21.41.00	Pasarsenen - Malang	EKONOMI	LANGSUNG
    287	MATARMAJA	15.15.00	Malang - Pasarsenen	EKONOMI	LANGSUNG
    288	MATARMAJA	20.22.00	Pasarsenen - Malang	EKONOMI	LANGSUNG
    337F	BANYUBIRU	09.28.00	Solobalapan - Semarang Tawang	EKONOMI	LANGSUNG
    338F	BANYUBIRU	08.12.00	Semarang Tawang - Solobalapan	EKONOMI	LANGSUNG
    2704F	BRAMNAMBO SERVICE	11.46.00	Nambo - Kampungbandan - Solobalapan - Brambanan	KA BARANG KHUSUS (SEMEN)	LANGSUNG
    2703F	BRAMNAMBO SERVICE	04.11.00	Brambanan - Solobalapan - Kampungbandan - Nambo	KA BARANG KHUSUS (SEMEN)	LANGSUNG
    2708F	BRAMARJA SERVICE	04.01.00	Arjawinangun - Solobalapan - Brambanan	KA BARANG KHUSUS (SEMEN)	LANGSUNG
    2707F	BRAMARJA SERVICE	13.29.00	Brambanan - Solobalapan - Arjawinangun	KA BARANG KHUSUS (SEMEN)	LANGSUNG
    2713	BUNGTALUN SERVICE	10.22.00	Karangtalun - Solobalapan - Brumbung	KA BARANG KHUSUS (SEMEN)	LANGSUNG
    2714	BUNGTALUN SERVICE	23.36.00	Brumbung - Solobalapan - Karangtalun	KA BARANG KHUSUS (SEMEN)	LANGSUNG
    `,
    // #DATABASE_STASIUN_SALEM
    SALEM: ``,
    // #DATABASE_STASIUN_SUMBERLAWANG
    SUMBERLAWANG: ``,
    // #DATABASE_STASIUN_GOPRAK
    GOPRAK: ``,
    // #DATABASE_STASIUN_JENAR
    JENAR: ``,
    // #DATABASE_STASIUN_WOJO
    WOJO: ``,
    // #DATABASE_STASIUN_SENTOLO
    SENTOLO: ``,
    // #DATABASE_STASIUN_BANDARA_YIA
    BANDARA_YIA: ``,
    // #DATABASE_STASIUN_KEMIRI
    KEMIRI: ``,
    // #DATABASE_STASIUN_MASARAN
    MASARAN: ``,
    // #DATABASE_STASIUN_SRAGEN
    SRAGEN: ``,
    // #DATABASE_STASIUN_KEBONROMO
    KEBONROMO: ``,
    // #DATABASE_STASIUN_KEDUNGBANTENG
    KEDUNGBANTENG: ``
};

/* ==========================================================================
   METADATA PETA JARINGAN KERETA API DAOP 6 YOGYAKARTA (36 STASIUN)
   ========================================================================== */
const DAOP6_NETWORK_NODES = [
    // --- LINTAS UTAMA BARAT KE TIMUR (y = 260) ---
    { key: 'JENAR',         name: 'Jenar',                code: 'JN',  kelas: 'Kelas III',            tier: 'kelas3', x: 70,   y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'WOJO',          name: 'Wojo',                 code: 'WJ',  kelas: 'Kelas III',            tier: 'kelas3', x: 125,  y: 260, lblPos: 'bottom', line: 'Lintas Utama Barat–Timur' },
    { key: 'KEDUNDANG',     name: 'Kedundang',            code: 'KDG', kelas: 'Belum teridentifikasi',tier: 'kelas3', x: 185,  y: 260, lblPos: 'top',    line: 'Simpul Percabangan Bandara YIA' },
    { key: 'WATES',         name: 'Wates',                code: 'WT',  kelas: 'Kelas II*',            tier: 'kelas2', x: 245,  y: 260, lblPos: 'bottom', line: 'Lintas Utama Barat–Timur' },
    { key: 'SENTOLO',       name: 'Sentolo',              code: 'STL', kelas: 'Kelas III',            tier: 'kelas3', x: 300,  y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'REWULU',        name: 'Rewulu',               code: 'RWL', kelas: 'Kelas I',              tier: 'kelas1', x: 355,  y: 260, lblPos: 'bottom', line: 'Lintas Utama (Depo BBM Pertamina)' },
    { key: 'PATUKAN',       name: 'Patukan',              code: 'PTN', kelas: 'Kelas III',            tier: 'kelas3', x: 410,  y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'YOGYAKARTA',    name: 'Yogyakarta (Tugu)',    code: 'YK',  kelas: 'Besar A',              tier: 'besar',  x: 475,  y: 260, lblPos: 'bottom', line: 'Lintas Utama • Terminus YIA & Prameks' },
    { key: 'LEMPUYANGAN',   name: 'Lempuyangan',          code: 'LPN', kelas: 'Besar B',              tier: 'besar',  x: 545,  y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'MAGUWO',        name: 'Maguwo',               code: 'MGW', kelas: 'Kelas II',             tier: 'kelas2', x: 605,  y: 260, lblPos: 'bottom', line: 'Lintas Utama • Integrasi Bandara Adisutjipto' },
    { key: 'BRAMBANAN',     name: 'Brambanan',            code: 'BBN', kelas: 'Kelas II*',            tier: 'kelas2', x: 665,  y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'SROWOT',        name: 'Srowot',               code: 'SWT', kelas: 'Kelas III',            tier: 'kelas3', x: 720,  y: 260, lblPos: 'bottom', line: 'Lintas Utama Barat–Timur' },
    { key: 'KLATEN',        name: 'Klaten',               code: 'KT',  kelas: 'Kelas I',              tier: 'kelas1', x: 775,  y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'CEPER',         name: 'Ceper',                code: 'CE',  kelas: 'Kelas II',             tier: 'kelas2', x: 830,  y: 260, lblPos: 'bottom', line: 'Lintas Utama Barat–Timur' },
    { key: 'DELANGGU',      name: 'Delanggu',             code: 'DL',  kelas: 'Kelas III',            tier: 'kelas3', x: 885,  y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'GAWOK',         name: 'Gawok',                code: 'GW',  kelas: 'Kelas III',            tier: 'kelas3', x: 940,  y: 260, lblPos: 'bottom', line: 'Lintas Utama Barat–Timur' },
    { key: 'PURWOSARI',     name: 'Purwosari',            code: 'PWS', kelas: 'Besar C',              tier: 'besar',  x: 1005, y: 260, lblPos: 'top',    line: 'Simpul Percabangan Jalur Wonogiri' },
    { key: 'SOLO_BALAPAN',  name: 'Solobalapan',          code: 'SLO', kelas: 'Besar A',              tier: 'besar',  x: 1085, y: 260, lblPos: 'bottom', line: 'Simpul Percabangan Utara (Gundih & Bandara SMO)' },
    { key: 'SOLOJEBRES',    name: 'Solojebres',           code: 'SK',  kelas: 'Besar C',              tier: 'besar',  x: 1160, y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'PALUR',         name: 'Palur',                code: 'PL',  kelas: 'Kelas III',            tier: 'kelas3', x: 1220, y: 260, lblPos: 'bottom', line: 'Lintas Utama • Terminus Timur KRL Jogja–Solo' },
    { key: 'KEMIRI',        name: 'Kemiri',               code: 'KMR', kelas: 'Kelas III',            tier: 'kelas3', x: 1275, y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'MASARAN',       name: 'Masaran',              code: 'MSR', kelas: 'Kelas III',            tier: 'kelas3', x: 1330, y: 260, lblPos: 'bottom', line: 'Lintas Utama Barat–Timur' },
    { key: 'SRAGEN',        name: 'Sragen',               code: 'SR',  kelas: 'Kelas II',             tier: 'kelas2', x: 1385, y: 260, lblPos: 'top',    line: 'Lintas Utama Barat–Timur' },
    { key: 'KEBONROMO',     name: 'Kebonromo',            code: 'KRO', kelas: 'Kelas III',            tier: 'kelas3', x: 1440, y: 260, lblPos: 'bottom', line: 'Lintas Utama Barat–Timur' },
    { key: 'KEDUNGBANTENG', name: 'Kedungbanteng',        code: 'KDB', kelas: 'Kelas III',            tier: 'kelas3', x: 1500, y: 260, lblPos: 'top',    line: 'Batas Timur Wilayah Daop 6 (Arah Madiun/Surabaya)' },

    // --- CABANG SELATAN DARI KEDUNDANG KE BANDARA YIA ---
    { key: 'BANDARA_YIA',   name: 'Bandara YIA',          code: 'YIA', kelas: 'Kelas I',              tier: 'kelas1', x: 185,  y: 355, lblPos: 'bottom', line: 'Cabang Selatan (KA Bandara Yogyakarta Int. Airport)' },

    // --- CABANG UTARA DARI SOLOBALAPAN KE GUNDIH ---
    { key: 'KADIPIRO',      name: 'Kadipiro',             code: 'KDO', kelas: 'Belum teridentifikasi',tier: 'kelas3', x: 1085, y: 195, lblPos: 'right',  line: 'Cabang Utara • Simpul Percabangan Bandara SMO' },
    { key: 'KALIOSO',       name: 'Kalioso',              code: 'KO',  kelas: 'Kelas III',            tier: 'kelas3', x: 1085, y: 155, lblPos: 'right',  line: 'Cabang Utara (Solobalapan – Gundih)' },
    { key: 'SALEM',         name: 'Salem',                code: 'SLM', kelas: 'Kelas III',            tier: 'kelas3', x: 1085, y: 118, lblPos: 'right',  line: 'Cabang Utara (Solobalapan – Gundih)' },
    { key: 'SUMBERLAWANG',  name: 'Sumberlawang',         code: 'SUM', kelas: 'Kelas III',            tier: 'kelas3', x: 1085, y: 81,  lblPos: 'right',  line: 'Cabang Utara (Solobalapan – Gundih)' },
    { key: 'GOPRAK',        name: 'Goprak',               code: 'GPK', kelas: 'Kelas III',            tier: 'kelas3', x: 1085, y: 44,  lblPos: 'right',  line: 'Batas Utara Daop 6 (Lanjut ke Gundih / Semarang)' },

    // --- CABANG BARAT DARI KADIPIRO KE BANDARA ADI SOEMARMO ---
    { key: 'ADI_SOEMARMO',  name: 'Bandara Adi Soemarmo', code: 'SMO', kelas: 'Kelas I',              tier: 'kelas1', x: 930,  y: 195, lblPos: 'top',    line: 'Cabang Barat dari Kadipiro (KA Bandara BIAS)' },

    // --- CABANG TENGGARA DARI PURWOSARI KE WONOGIRI ---
    { key: 'SOLOKOTA',      name: 'Solokota',             code: 'STA', kelas: 'Kelas III',            tier: 'kelas3', x: 1045, y: 315, lblPos: 'right',  line: 'Cabang Tenggara (Purwosari – Wonogiri)' },
    { key: 'SUKOHARJO',     name: 'Sukoharjo',            code: 'SKH', kelas: 'Kelas III',            tier: 'kelas3', x: 1085, y: 360, lblPos: 'right',  line: 'Cabang Tenggara (Purwosari – Wonogiri)' },
    { key: 'PASARNGUTER',   name: 'Pasarnguter',          code: 'PNT', kelas: 'Kelas III',            tier: 'kelas3', x: 1125, y: 405, lblPos: 'right',  line: 'Cabang Tenggara (Purwosari – Wonogiri)' },
    { key: 'WONOGIRI',      name: 'Wonogiri',             code: 'WNG', kelas: 'Kelas III',            tier: 'kelas3', x: 1165, y: 450, lblPos: 'right',  line: 'Terminus Tenggara Daop 6 (Batara Kresna)' }
];

/* ==========================================================================
   PARSER EXCEL -> OBJEK JAVASCRIPT
   ========================================================================== */
function parseStationRawData(rawText) {
    if (!rawText || !rawText.trim()) return [];

    const lines = rawText.trim().split(/\r?\n/);
    const parsedList = [];

    lines.forEach(line => {
        const trimmed = line.trim();
        if (!trimmed) return;
        if (/^no\.?\s*ka/i.test(trimmed)) return;

        let cols = trimmed.split('\t').map(c => c.trim());
        if (cols.length < 5) {
            cols = trimmed.split(/\s{2,}/).map(c => c.trim());
        }

        if (cols.length >= 5) {
            const noKa = cols[0];
            const namaKa = cols[1];
            const rawTime = cols[2].replace(/\./g, ':');

            let relasi = '—';
            let kelas = '';
            let ketRaw = '';

            if (cols.length >= 6) {
                relasi = cols[3] || '—';
                kelas = cols[4];
                ketRaw = cols[5].toUpperCase();
            } else {
                kelas = cols[3];
                ketRaw = cols[4].toUpperCase();
            }

            const timeParts = rawTime.split(':').map(Number);
            const hours = timeParts[0] || 0;
            const minutes = timeParts[1] || 0;
            const seconds = timeParts[2] || 0;
            const totalSeconds = (hours * 3600) + (minutes * 60) + seconds;

            const formattedTime = [
                String(hours).padStart(2, '0'),
                String(minutes).padStart(2, '0'),
                String(seconds).padStart(2, '0')
            ].join(':');

            let statusKey = 'BERHENTI';
            let statusLabel = 'BERHENTI';
            if (ketRaw.includes('LANGSUNG')) {
                statusKey = 'LANGSUNG';
                statusLabel = 'BERJALAN LANGSUNG';
            } else if (ketRaw.includes('START') || ketRaw.includes('AWAL')) {
                statusKey = 'START';
                statusLabel = 'BERANGKAT AWAL (START)';
            } else if (ketRaw.includes('TERMINUS')) {
                statusKey = 'TERMINUS';
                statusLabel = 'TERMINUS';
            }

            const numericPart = parseInt(noKa.replace(/\D/g, ''), 10);
            const isOdd = !isNaN(numericPart) && (numericPart % 2 !== 0);

            parsedList.push({
                noKa,
                namaKa,
                time: formattedTime,
                totalSeconds,
                relasi,
                kelas,
                statusKey,
                statusLabel,
                isOdd
            });
        }
    });

    return parsedList.sort((a, b) => a.totalSeconds - b.totalSeconds);
}

const PARSED_DATABASE = {};
Object.keys(STATION_META).forEach(key => {
    PARSED_DATABASE[key] = parseStationRawData(RAW_STATION_DATABASE[key] || '');
});

/* ==========================================================================
   STATE & DOM ELEMENTS (DENGAN PERSISTENSI STASIUN VIA LOCALSTORAGE)
   ========================================================================== */
const savedStation = localStorage.getItem('selected_station_daop6');
let currentStation = (savedStation && STATION_META[savedStation]) ? savedStation : 'SOLO_BALAPAN';
let currentStatusFilter = 'ALL';
let currentSearchQuery = '';

function saveSelectedStation(stationKey) {
    currentStation = stationKey;
    try {
        localStorage.setItem('selected_station_daop6', stationKey);
    } catch (e) {
        // Abaikan jika browser memblokir storage (mode incognito ketat)
    }
}

const hamburgerBtn = document.getElementById('hamburger-btn');
const navDrawer = document.getElementById('nav-drawer');
const menuOverlay = document.getElementById('menu-overlay');
const closeDrawerBtn = document.getElementById('close-drawer');
const navItems = document.querySelectorAll('.nav-item[data-target]');
const externalNavItems = document.querySelectorAll('a.nav-item');
const pageSections = document.querySelectorAll('.page-section');

const rtStationName = document.getElementById('rt-station-name');
const rtCardsContainer = document.getElementById('realtime-cards-container');
const scheduleTbody = document.getElementById('schedule-tbody');
const emptyState = document.getElementById('empty-state');
const searchKaInput = document.getElementById('search-ka');
const filterPills = document.querySelectorAll('#filter-pills .pill');

/* ==========================================================================
   NAVIGASI HAMBURGER MENU & PERPINDAHAN HALAMAN (DENGAN URL HASH)
   ========================================================================== */
function toggleDrawer(open) {
    if (!navDrawer || !menuOverlay || !hamburgerBtn) return;
    if (open) {
        navDrawer.classList.add('active');
        menuOverlay.classList.add('active');
        hamburgerBtn.classList.add('open');
        hamburgerBtn.setAttribute('aria-expanded', 'true');
    } else {
        navDrawer.classList.remove('active');
        menuOverlay.classList.remove('active');
        hamburgerBtn.classList.remove('open');
        hamburgerBtn.setAttribute('aria-expanded', 'false');
    }
}

if (hamburgerBtn) {
    hamburgerBtn.addEventListener('click', () => {
        const isOpen = navDrawer.classList.contains('active');
        toggleDrawer(!isOpen);
    });
}

if (closeDrawerBtn) closeDrawerBtn.addEventListener('click', () => toggleDrawer(false));
if (menuOverlay) menuOverlay.addEventListener('click', () => toggleDrawer(false));

function switchPage(targetId, updateHash = true) {
    if (!targetId) return;

    const targetSection = document.getElementById(`section-${targetId}`);
    if (!targetSection) targetId = 'home';

    navItems.forEach(item => {
        item.classList.toggle('active', item.dataset.target === targetId);
    });

    pageSections.forEach(sec => {
        sec.classList.toggle('active', sec.id === `section-${targetId}`);
    });

    if (updateHash && window.location.hash !== `#${targetId}`) {
        history.replaceState(null, '', `#${targetId}`);
    }

    toggleDrawer(false);
    updateMiniPlayerVisibility();
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

navItems.forEach(item => {
    item.addEventListener('click', () => switchPage(item.dataset.target));
});

externalNavItems.forEach(link => {
    link.addEventListener('click', () => toggleDrawer(false));
});

const btnGoJadwal = document.getElementById('btn-go-jadwal');
if (btnGoJadwal) {
    btnGoJadwal.addEventListener('click', () => switchPage('jadwal'));
}

/* ==========================================================================
   MODUL PLAYLIST BEL & LAGU KEDATANGAN STASIUN (PERSISTENT BACKGROUND AUDIO)
   ========================================================================== */
const STATION_AUDIO_PLAYLIST = [
    {
        id: 'bel-westminster',
        title: 'Westminster Chimes (Bel Pengumuman)',
        station: 'Bel Standar Stasiun Kereta Api',
        category: 'BEL',
        src: 'audio/bel-stasiun.mp3'
    },
    {
        id: 'lagu-bengawan-solo',
        title: 'Bengawan Solo (Instrumental)',
        station: 'Stasiun Solo Balapan (SLO) & Purwosari (PWS)',
        category: 'LAGU',
        src: 'audio/bengawan-solo.mp3'
    },
    {
        id: 'lagu-sepasang-mata-bola',
        title: 'Sepasang Mata Bola (Instrumental)',
        station: 'Stasiun Yogyakarta (YK)',
        category: 'LAGU',
        src: 'audio/sepasang-mata-bola.mp3'
    },
    {
        id: 'lagu-sepasang-mata-bola',
        title: 'Sepasang Mata Bola (Instrumental)',
        station: 'Stasiun Lempuyangan (LPN)',
        category: 'LAGU',
        src: 'audio/sepasang-mata-bola-lpn.mp3'
    }
];

const stationAudio = document.getElementById('station-audio-player');
const npDisc = document.getElementById('np-disc');
const npEqBars = document.getElementById('np-eq-bars');
const npStatusDot = document.getElementById('np-status-dot');
const npStateLabel = document.getElementById('np-state-label');
const npCategory = document.getElementById('np-category');
const npTitle = document.getElementById('np-title');
const npSubtitle = document.getElementById('np-subtitle');
const npSeekBar = document.getElementById('np-seek-bar');
const npCurrentTime = document.getElementById('np-current-time');
const npDuration = document.getElementById('np-duration');
const btnAudioPlay = document.getElementById('btn-audio-play');
const btnAudioPrev = document.getElementById('btn-audio-prev');
const btnAudioNext = document.getElementById('btn-audio-next');
const btnAudioLoop = document.getElementById('btn-audio-loop');
const btnAudioMute = document.getElementById('btn-audio-mute');
const npVolumeSlider = document.getElementById('np-volume-slider');
const npVolumeText = document.getElementById('np-volume-text');

const playlistContainer = document.getElementById('station-playlist-container');
const playlistSearchInput = document.getElementById('playlist-search-input');
const playlistTabs = document.querySelectorAll('#playlist-filter-tabs .pl-tab');

const floatingMiniPlayer = document.getElementById('floating-mini-player');
const miniTitle = document.getElementById('mini-title');
const miniSubtitle = document.getElementById('mini-subtitle');
const miniEq = document.getElementById('mini-eq');
const miniBtnPlay = document.getElementById('mini-btn-play');
const miniBtnNext = document.getElementById('mini-btn-next');
const miniOpenStudio = document.getElementById('mini-open-studio');

let currentTrackIndex = 0;
let playlistCategoryFilter = 'ALL';
let playlistSearchQuery = '';
let loopMode = 'ALL';

function formatAudioSeconds(sec) {
    if (!isFinite(sec) || sec < 0) return '00:00';
    const m = Math.floor(sec / 60);
    const s = Math.floor(sec % 60);
    return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

function updateMiniPlayerVisibility() {
    if (!floatingMiniPlayer || !stationAudio) return;
    const isAudioPageActive = document.getElementById('section-audio')?.classList.contains('active');
    const isPlayingOrPausedMidTrack = !stationAudio.paused || stationAudio.currentTime > 0;

    if (!isAudioPageActive && isPlayingOrPausedMidTrack) {
        floatingMiniPlayer.classList.remove('hidden');
    } else {
        floatingMiniPlayer.classList.add('hidden');
    }
}

function syncAudioPlayingUI(isPlaying, customErrorMsg = null) {
    if (isPlaying) {
        btnAudioPlay.textContent = '⏸';
        miniBtnPlay.textContent = '⏸';
        npDisc.classList.add('spinning');
        npEqBars.classList.add('active');
        miniEq.classList.add('active');
        npStatusDot.classList.add('live');
        npStateLabel.textContent = 'SEDANG DIPUTAR';
        miniSubtitle.textContent = 'Sedang diputar • Klik untuk buka playlist';
    } else {
        btnAudioPlay.textContent = '▶';
        miniBtnPlay.textContent = '▶';
        npDisc.classList.remove('spinning');
        npEqBars.classList.remove('active');
        miniEq.classList.remove('active');
        npStatusDot.classList.remove('live');
        npStateLabel.textContent = customErrorMsg ? 'FILE AUDIO TIDAK DITEMUKAN' : 'DIJEDA';
        miniSubtitle.textContent = 'Dijeda • Klik untuk buka playlist';
    }
    renderAudioPlaylist();
    updateMiniPlayerVisibility();
}

function loadAudioTrack(index, autoPlay = false) {
    if (STATION_AUDIO_PLAYLIST.length === 0) return;
    currentTrackIndex = (index + STATION_AUDIO_PLAYLIST.length) % STATION_AUDIO_PLAYLIST.length;
    const track = STATION_AUDIO_PLAYLIST[currentTrackIndex];

    stationAudio.src = track.src;
    npCategory.textContent = track.category === 'BEL' ? 'BEL PERON' : 'LAGU KEDATANGAN';
    npTitle.textContent = track.title;
    npSubtitle.textContent = track.station;
    miniTitle.textContent = track.title;

    npSeekBar.value = 0;
    npCurrentTime.textContent = '00:00';
    npDuration.textContent = '00:00';

    if (autoPlay) {
        stationAudio.play()
            .then(() => syncAudioPlayingUI(true))
            .catch(() => {
                syncAudioPlayingUI(false, true);
                npSubtitle.textContent = `Pastikan file "${track.src}" tersedia di folder lokal.`;
            });
    } else {
        syncAudioPlayingUI(false);
    }
}

function togglePlayPauseAudio() {
    if (!stationAudio.src) {
        loadAudioTrack(currentTrackIndex, true);
        return;
    }

    if (stationAudio.paused) {
        const track = STATION_AUDIO_PLAYLIST[currentTrackIndex];
        stationAudio.play()
            .then(() => syncAudioPlayingUI(true))
            .catch(() => {
                syncAudioPlayingUI(false, true);
                npSubtitle.textContent = `Pastikan file "${track.src}" tersedia di folder lokal.`;
            });
    } else {
        stationAudio.pause();
        syncAudioPlayingUI(false);
    }
}

function renderAudioPlaylist() {
    if (!playlistContainer) return;

    const countAllEl = document.getElementById('pl-count-all');
    if (countAllEl) countAllEl.textContent = STATION_AUDIO_PLAYLIST.length;

    const filtered = STATION_AUDIO_PLAYLIST
        .map((item, originalIndex) => ({ ...item, originalIndex }))
        .filter(item => {
            const matchCat = playlistCategoryFilter === 'ALL' || item.category === playlistCategoryFilter;
            const q = playlistSearchQuery.toLowerCase();
            const matchQuery = !q ||
                item.title.toLowerCase().includes(q) ||
                item.station.toLowerCase().includes(q);
            return matchCat && matchQuery;
        });

    if (filtered.length === 0) {
        playlistContainer.innerHTML = `<div class="pl-empty">Tidak ditemukan lagu atau bel yang cocok dengan pencarian.</div>`;
        return;
    }

    playlistContainer.innerHTML = filtered.map((track, idx) => {
        const isCurrent = track.originalIndex === currentTrackIndex;
        const isActivelyPlaying = isCurrent && !stationAudio.paused;

        return `
            <div class="pl-item ${isCurrent ? 'active' : ''}" data-index="${track.originalIndex}">
                <div class="pl-item-num">
                    ${isActivelyPlaying
                        ? `<span class="pl-playing-wave"><span></span><span></span><span></span></span>`
                        : `<span>${String(idx + 1).padStart(2, '0')}</span>`
                    }
                </div>
                <div class="pl-item-info">
                    <strong>${track.title}</strong>
                    <small>${track.station}</small>
                </div>
                <span class="pl-item-tag ${track.category === 'BEL' ? 'tag-bel' : 'tag-lagu'}">
                    ${track.category === 'BEL' ? 'BEL' : 'LAGU'}
                </span>
            </div>
        `;
    }).join('');

    playlistContainer.querySelectorAll('.pl-item').forEach(row => {
        row.addEventListener('click', () => {
            const idx = parseInt(row.dataset.index, 10);
            if (idx === currentTrackIndex) {
                togglePlayPauseAudio();
            } else {
                loadAudioTrack(idx, true);
            }
        });
    });
}

if (stationAudio) {
    stationAudio.volume = 0.8;

    btnAudioPlay.addEventListener('click', togglePlayPauseAudio);
    miniBtnPlay.addEventListener('click', togglePlayPauseAudio);

    btnAudioPrev.addEventListener('click', () => loadAudioTrack(currentTrackIndex - 1, true));
    btnAudioNext.addEventListener('click', () => loadAudioTrack(currentTrackIndex + 1, true));
    miniBtnNext.addEventListener('click', () => loadAudioTrack(currentTrackIndex + 1, true));

    miniOpenStudio.addEventListener('click', () => switchPage('audio'));

    btnAudioLoop.addEventListener('click', () => {
        if (loopMode === 'ALL') {
            loopMode = 'ONE';
            btnAudioLoop.textContent = '🔂';
            btnAudioLoop.classList.add('active-loop');
            btnAudioLoop.title = 'Mode Ulangi: 1 Lagu Terus-Menerus';
        } else {
            loopMode = 'ALL';
            btnAudioLoop.textContent = '🔁';
            btnAudioLoop.classList.remove('active-loop');
            btnAudioLoop.title = 'Mode Ulangi: Semua Playlist';
        }
    });

    stationAudio.addEventListener('loadedmetadata', () => {
        npDuration.textContent = formatAudioSeconds(stationAudio.duration);
    });

    stationAudio.addEventListener('timeupdate', () => {
        npCurrentTime.textContent = formatAudioSeconds(stationAudio.currentTime);
        if (stationAudio.duration) {
            npSeekBar.value = (stationAudio.currentTime / stationAudio.duration) * 100;
        }
    });

    npSeekBar.addEventListener('input', () => {
        if (stationAudio.duration) {
            stationAudio.currentTime = (npSeekBar.value / 100) * stationAudio.duration;
        }
    });

    stationAudio.addEventListener('ended', () => {
        if (loopMode === 'ONE') {
            stationAudio.currentTime = 0;
            stationAudio.play();
        } else {
            loadAudioTrack(currentTrackIndex + 1, true);
        }
    });

    npVolumeSlider.addEventListener('input', () => {
        const val = parseInt(npVolumeSlider.value, 10);
        stationAudio.volume = val / 100;
        stationAudio.muted = val === 0;
        npVolumeText.textContent = `${val}%`;
        btnAudioMute.textContent = val === 0 ? '🔇' : val < 50 ? '🔉' : '🔊';
    });

    btnAudioMute.addEventListener('click', () => {
        stationAudio.muted = !stationAudio.muted;
        if (stationAudio.muted) {
            btnAudioMute.textContent = '🔇';
            npVolumeText.textContent = 'Mute';
        } else {
            const val = parseInt(npVolumeSlider.value, 10) || 80;
            stationAudio.volume = val / 100;
            btnAudioMute.textContent = val < 50 ? '🔉' : '🔊';
            npVolumeText.textContent = `${val}%`;
        }
    });

    playlistSearchInput.addEventListener('input', (e) => {
        playlistSearchQuery = e.target.value;
        renderAudioPlaylist();
    });

    playlistTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            playlistTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            playlistCategoryFilter = tab.dataset.cat;
            renderAudioPlaylist();
        });
    });

    loadAudioTrack(0, false);
}

function playTrainHorn() {
    const s35Audio = document.getElementById('s35-audio-player');
    if (!s35Audio) return;

    s35Audio.currentTime = 0;
    s35Audio.volume = 0.9;
    s35Audio.play().catch(err => {
        console.warn('File audio Semboyan 35 belum ditemukan atau diblokir browser:', err);
    });
}

const btnTriggerHorn = document.getElementById('btn-trigger-horn');
const interactiveTrain = document.getElementById('interactive-train');
if (btnTriggerHorn) btnTriggerHorn.addEventListener('click', playTrainHorn);
if (interactiveTrain) interactiveTrain.addEventListener('click', playTrainHorn);

/* ==========================================================================
   LOGIKA ARAH PERJALANAN KA (LINTAS UTAMA BARAT-TIMUR VS CABANG UTARA GUNDIH)
   ========================================================================== */
const NORTH_BRANCH_STATIONS = ['KADIPIRO', 'KALIOSO', 'SALEM', 'SUMBERLAWANG', 'GOPRAK'];

function isNorthBranchStation(stationKey) {
    return NORTH_BRANCH_STATIONS.includes(stationKey);
}

function getDirectionText(isOdd) {
    if (isNorthBranchStation(currentStation)) {
        return isOdd ? 'Selatan → Utara' : 'Utara → Selatan';
    }
    return isOdd ? 'Timur → Barat' : 'Barat → Timur';
}

function renderDirectionArrow(isOdd) {
    if (isNorthBranchStation(currentStation)) {
        const upSvg = `<svg viewBox="0 0 12 12"><polyline points="2 8 6 4 10 8"></polyline></svg>`;
        const downSvg = `<svg viewBox="0 0 12 12"><polyline points="2 4 6 8 10 4"></polyline></svg>`;

        if (isOdd) {
            return `<span class="dir-arrow-pill odd-up" title="KA Ganjil — Arah Utara (Selatan ke Utara / Gundih)">${upSvg}${upSvg}${upSvg}</span>`;
        }
        return `<span class="dir-arrow-pill even-down" title="KA Genap — Arah Selatan (Utara ke Selatan / Solo Balapan)">${downSvg}${downSvg}${downSvg}</span>`;
    }

    const leftSvg = `<svg viewBox="0 0 12 12"><polyline points="8 2 4 6 8 10"></polyline></svg>`;
    const rightSvg = `<svg viewBox="0 0 12 12"><polyline points="4 2 8 6 4 10"></polyline></svg>`;

    if (isOdd) {
        return `<span class="dir-arrow-pill odd-left" title="KA Ganjil — Arah Kiri (Timur ke Barat)">${leftSvg}${leftSvg}${leftSvg}</span>`;
    }
    return `<span class="dir-arrow-pill even-right" title="KA Genap — Arah Kanan (Barat ke Timur)">${rightSvg}${rightSvg}${rightSvg}</span>`;
}

/* ==========================================================================
   KALKULASI RADAR 4 ZONA OPERASIONAL & RENDER JADWAL
   ========================================================================== */
function getTrainPresenceConfig(train) {
    const status = train.statusKey;
    const kelasUpper = (train.kelas || '').toUpperCase();
    const namaUpper = (train.namaKa || '').toUpperCase();

    if (status === 'LANGSUNG') {
        return {
            preActiveSeconds: 0,
            totalActiveSeconds: 60,
            phase: 'PASSING',
            badgeText: (elapsed) => `Melintas Ls • ${Math.max(1, 60 - elapsed)} dtk`,
            zoneClass: 'is-passing'
        };
    }

    if (status === 'START') {
        const preBoardSec = 600;
        return {
            preActiveSeconds: preBoardSec,
            totalActiveSeconds: 0,
            phase: 'START_BOARDING',
            badgeText: (elapsed) => {
                const remToDepart = Math.max(0, -elapsed);
                const m = Math.floor(remToDepart / 60);
                const s = remToDepart % 60;
                return m > 0 ? `Standby Start • berangkat ${m}m ${s}d` : `Siap Berangkat • ${s}d`;
            },
            zoneClass: 'is-start'
        };
    }

    if (status === 'BERHENTI') {
        const isArgo = kelasUpper.includes('ARGO');
        const dwellSec = isArgo ? 180 : 300;
        return {
            preActiveSeconds: 0,
            totalActiveSeconds: dwellSec,
            phase: 'PLATFORM',
            badgeText: (elapsed) => {
                const rem = Math.max(0, dwellSec - elapsed);
                const m = Math.floor(rem / 60);
                const s = rem % 60;
                const tag = isArgo ? 'Stop 3 mnt' : 'Stop 5 mnt';
                return m > 0 ? `${tag} • sisa ${m}m ${s}d` : `Siap berangkat • ${s}d`;
            },
            zoneClass: 'is-platform'
        };
    }

    const isCommuterOrAirport = kelasUpper.includes('KRL') ||
                                kelasUpper.includes('BANDARA') ||
                                kelasUpper.includes('KRD') ||
                                namaUpper.includes('COMMUTER') ||
                                namaUpper.includes('PRAMEKS');

    if (isCommuterOrAirport) {
        const dwellSec = 480;
        return {
            preActiveSeconds: 0,
            totalActiveSeconds: dwellSec,
            phase: 'PLATFORM',
            badgeText: (elapsed) => {
                const rem = Math.max(0, dwellSec - elapsed);
                const m = Math.floor(rem / 60);
                return `Terminus Peron • ±${m} mnt`;
            },
            zoneClass: 'is-platform'
        };
    }

    const platformEndSec = 600;
    const depoEndSec = 1500;

    return {
        preActiveSeconds: 0,
        totalActiveSeconds: depoEndSec,
        phase: 'TERMINUS_DYNAMIC',
        isInDepo: (elapsed) => elapsed > platformEndSec,
        badgeText: (elapsed) => {
            if (elapsed <= platformEndSec) {
                const rem = platformEndSec - elapsed;
                const m = Math.floor(rem / 60);
                return `Di Peron Terminus • ±${m} mnt`;
            }
            const depoMins = Math.floor(elapsed / 60);
            return `Stabling / Depo KA (${depoMins} mnt)`;
        },
        zoneClass: 'is-terminus'
    };
}

function getCurrentSecondsOfDay() {
    const now = new Date();
    return (now.getHours() * 3600) + (now.getMinutes() * 60) + now.getSeconds();
}

function formatDuration(diffSeconds) {
    if (diffSeconds <= 0) return "TIBA SEKARANG";
    const hrs = Math.floor(diffSeconds / 3600);
    const mins = Math.floor((diffSeconds % 3600) / 60);
    const secs = diffSeconds % 60;
    if (hrs > 0) return `${hrs} jam ${mins} mnt lagi`;
    if (mins > 0) return `${mins} mnt ${secs} dtk lagi`;
    return `${secs} dtk lagi`;
}

function formatPastDuration(pastSeconds) {
    if (pastSeconds < 60) return `${pastSeconds} dtk lalu`;
    const hrs = Math.floor(pastSeconds / 3600);
    const mins = Math.floor((pastSeconds % 3600) / 60);
    if (hrs > 0) return `${hrs} jam ${mins} mnt lalu`;
    return `${mins} mnt lalu`;
}

function getBadgeClass(statusKey) {
    if (statusKey === 'LANGSUNG') return 'status-langsung';
    if (statusKey === 'START') return 'status-start';
    if (statusKey === 'TERMINUS') return 'status-terminus';
    return 'status-berhenti';
}

function renderTrainSubBlock(train, badgeLabel, extraNote = '', customStatePill = null) {
    const statusBadgeHtml = customStatePill
        ? customStatePill
        : `<span class="status-badge ${getBadgeClass(train.statusKey)}">${train.statusLabel}</span>`;

    return `
        <div class="rt-sub-item">
            <div class="rt-card-top">
                <span class="rt-badge-time">${train.time} WIB</span>
                <span class="rt-countdown">${badgeLabel}</span>
            </div>
            <div>
                <div class="rt-train-name">
                    <span>KA ${train.noKa} — ${train.namaKa}</span>
                    ${renderDirectionArrow(train.isOdd)}
                </div>
                ${train.relasi !== '—' ? `<div class="rt-train-relasi">Relasi: ${train.relasi}</div>` : ''}
                <div class="rt-train-meta">Kelas: ${train.kelas} ${extraNote}</div>
            </div>
            <div class="rt-card-bottom">
                ${statusBadgeHtml}
                <span class="rt-dir-label">${getDirectionText(train.isOdd)}</span>
            </div>
        </div>
    `;
}

function renderRealtimeRadar() {
    if (!rtStationName || !rtCardsContainer) return null;

    const data = PARSED_DATABASE[currentStation] || [];
    const meta = STATION_META[currentStation];
    rtStationName.textContent = `${meta.name.toUpperCase()} (${meta.code})`;

    const rtNoteEl = document.querySelector('.rt-update-note');
    if (rtNoteEl) {
        rtNoteEl.textContent = isNorthBranchStation(currentStation)
            ? 'Estimasi GAPEKA 2026 • ↑ Ganjil (Utara) | ↓ Genap (Selatan)'
            : 'Estimasi GAPEKA 2026 • « Ganjil (Barat) | Genap (Timur) »';
    }

    if (data.length === 0) {
        rtCardsContainer.innerHTML = `
            <div class="rt-card">
                <div class="rt-train-name">Database Stasiun Belum Tersedia</div>
                <div class="rt-train-meta">Mohon maaf jadwal untuk ${meta.name} belum tersedia</div>
            </div>
        `;
        return null;
    }

    const nowSec = getCurrentSecondsOfDay();

    const enriched = data.map(item => {
        let elapsed = nowSec - item.totalSeconds;
        let isYesterday = false;

        if (elapsed < -84000) {
            elapsed = nowSec + (86400 - item.totalSeconds);
            isYesterday = true;
        }

        const presence = getTrainPresenceConfig(item);
        return { ...item, elapsed, isYesterday, presence };
    });

    // ZONA 2 (KOTAK 2): Sedang di Stasiun / Melintas Emplasemen / Standby Start
    const atStationList = enriched
        .filter(t => {
            if (t.statusKey === 'START') {
                return t.elapsed >= -t.presence.preActiveSeconds && t.elapsed < 0;
            }
            return t.elapsed >= 0 && t.elapsed <= t.presence.totalActiveSeconds;
        })
        .sort((a, b) => a.elapsed - b.elapsed);

    // ZONA 1 (KOTAK 1): 1 Kereta Terakhir Berlalu
    const passedCandidates = enriched
        .filter(t => t.elapsed > t.presence.totalActiveSeconds)
        .sort((a, b) => a.elapsed - b.elapsed);

    let lastPassed = passedCandidates[0] || null;
    if (!lastPassed && data.length > 0) {
        const lastNight = data[data.length - 1];
        lastPassed = {
            ...lastNight,
            elapsed: nowSec + (86400 - lastNight.totalSeconds),
            isYesterday: true
        };
    }

    // ZONA 3 & 4: Kereta yang belum tiba
    let futureList = data
        .filter(t => t.totalSeconds > nowSec)
        .map(t => ({ ...t, diff: t.totalSeconds - nowSec, isTomorrow: false }));

    if (futureList.length < 4) {
        const needed = 4 - futureList.length;
        for (let i = 0; i < Math.min(needed, data.length); i++) {
            futureList.push({
                ...data[i],
                diff: (86400 - nowSec) + data[i].totalSeconds,
                isTomorrow: true
            });
        }
    }

    let nextGroup = [];
    let queueStartIndex = 0;
    if (futureList.length > 0) {
        const nearestTime = futureList[0].diff;
        nextGroup = futureList.filter(t => t.diff === nearestTime);
        queueStartIndex = nextGroup.length;
    }

    const queueList = futureList.slice(queueStartIndex, queueStartIndex + 2);

    const box1Html = `
        <div class="rt-card zone-past">
            <div class="rt-zone-header">
                <span class="rt-zone-tag">1 • TERAKHIR BERLALU</span>
            </div>
            <div class="rt-stack-list">
                ${lastPassed
                    ? renderTrainSubBlock(
                        lastPassed,
                        formatPastDuration(lastPassed.elapsed),
                        lastPassed.isYesterday ? '(Kemarin)' : ''
                      )
                    : `<div class="rt-empty-slot">Belum ada perjalanan berlalu.</div>`
                }
            </div>
        </div>
    `;

    const box2Html = `
        <div class="rt-card zone-at-station ${atStationList.length === 0 ? 'is-idle' : ''}">
            <div class="rt-zone-header">
                <span class="rt-zone-tag">2 • AKTIVITAS DI STASIUN</span>
                ${atStationList.length > 1 ? `<span class="rt-zone-count">${atStationList.length} KA AKTIF</span>` : ''}
            </div>
            <div class="rt-stack-list">
                ${atStationList.length > 0
                    ? atStationList.map(t => {
                        const dynamicBadgeText = t.presence.badgeText(t.elapsed);
                        let customPill = null;

                        if (t.statusKey === 'LANGSUNG') {
                            customPill = `<span class="status-badge status-passing-live">⚡ MELINTAS LANGSUNG</span>`;
                        } else if (t.statusKey === 'START') {
                            customPill = `<span class="status-badge status-start-live">🚆 STANDBY KEBERANGKATAN</span>`;
                        } else if (t.presence.isInDepo && t.presence.isInDepo(t.elapsed)) {
                            customPill = `<span class="status-badge status-depo-live">🔁 LANGSIR / DI DEPO KA</span>`;
                        }

                        return renderTrainSubBlock(
                            t,
                            dynamicBadgeText,
                            t.isYesterday ? '(Kemarin)' : '',
                            customPill
                        );
                    }).join('')
                    : `<div class="rt-empty-slot">Emplasemen & peron kosong • Tidak ada KA berhenti (3–5 mnt), melintas (1 mnt), atau terminus aktif saat ini.</div>`
                }
            </div>
        </div>
    `;

    const box3Html = `
        <div class="rt-card zone-next">
            <div class="rt-zone-header">
                <span class="rt-zone-tag">3 • TERDEKAT BERIKUTNYA</span>
                ${nextGroup.length > 1 ? `<span class="rt-zone-count">${nextGroup.length} KA BERSAMAAN</span>` : ''}
            </div>
            <div class="rt-stack-list">
                ${nextGroup.length > 0
                    ? nextGroup.map(t => renderTrainSubBlock(
                        t,
                        formatDuration(t.diff),
                        t.isTomorrow ? '(Besok)' : ''
                      )).join('')
                    : `<div class="rt-empty-slot">Tidak ada jadwal berikutnya.</div>`
                }
            </div>
        </div>
    `;

    const box4Html = `
        <div class="rt-card zone-queue">
            <div class="rt-zone-header">
                <span class="rt-zone-tag">4 • ANTRIAN SELANJUTNYA</span>
            </div>
            <div class="rt-stack-list">
                ${queueList.length > 0
                    ? queueList.map((t, idx) => `
                        <div class="rt-queue-item">
                            <div class="rt-queue-top">
                                <span class="rt-queue-time">${t.time} WIB</span>
                                <span class="rt-countdown">#${idx + 2} • ${formatDuration(t.diff)}</span>
                            </div>
                            <div class="rt-queue-title">
                                <span>KA ${t.noKa} — ${t.namaKa}</span>
                                ${renderDirectionArrow(t.isOdd)}
                            </div>
                            <div class="rt-queue-sub">
                                <span>${t.relasi !== '—' ? t.relasi : t.kelas}</span>
                                <span class="status-badge ${getBadgeClass(t.statusKey)}">${t.statusKey}</span>
                            </div>
                        </div>
                    `).join('')
                    : `<div class="rt-empty-slot">Tidak ada antrian berikutnya.</div>`
                }
            </div>
        </div>
    `;

    rtCardsContainer.innerHTML = box1Html + box2Html + box3Html + box4Html;

    return nextGroup[0] ? `${nextGroup[0].noKa}-${nextGroup[0].time}` : null;
}

function renderScheduleTable(nextTrainKey) {
    if (!scheduleTbody || !emptyState) return;
    const data = PARSED_DATABASE[currentStation] || [];

    const filtered = data.filter(item => {
        const matchStatus = currentStatusFilter === 'ALL' || item.statusKey === currentStatusFilter;
        const q = currentSearchQuery.toLowerCase();
        const matchQuery = !q ||
            item.noKa.toLowerCase().includes(q) ||
            item.namaKa.toLowerCase().includes(q) ||
            item.relasi.toLowerCase().includes(q) ||
            item.kelas.toLowerCase().includes(q);
        return matchStatus && matchQuery;
    });

    if (filtered.length === 0) {
        scheduleTbody.innerHTML = '';
        emptyState.classList.remove('hidden');
        return;
    }

    emptyState.classList.add('hidden');
    scheduleTbody.innerHTML = filtered.map(item => {
        const rowKey = `${item.noKa}-${item.time}`;
        const isNext = rowKey === nextTrainKey;
        const dirHint = isNorthBranchStation(currentStation)
            ? (item.isOdd ? '↑' : '↓')
            : (item.isOdd ? '«' : '»');
        return `
            <tr class="${isNext ? 'is-next-row' : ''}">
                <td class="col-time">${item.time}</td>
                <td class="col-ka">${item.noKa} <span style="opacity:0.65;">${dirHint}</span></td>
                <td class="col-name">${item.namaKa} ${isNext ? '⚡' : ''}</td>
                <td class="col-relasi">${item.relasi}</td>
                <td class="col-class">${item.kelas}</td>
                <td><span class="status-badge ${getBadgeClass(item.statusKey)}">${item.statusLabel}</span></td>
            </tr>
        `;
    }).join('');
}

function updateFullScheduleView() {
    const nextTrainKey = renderRealtimeRadar();
    renderScheduleTable(nextTrainKey);
}

/* ==========================================================================
   HELPER DATA STASIUN TERURUT ABJAD & KONTROL CUSTOM DROPDOWN / DIREKTORI
   ========================================================================== */
const CLASS_RANK_MAP = {
    'Besar A': 1,
    'Besar B': 2,
    'Besar C': 3,
    'Kelas I': 4,
    'Kelas II*': 5,
    'Kelas II': 6,
    'Kelas III': 7,
    'Belum teridentifikasi': 8
};

function getAllStationsList() {
    const nodeMap = {};
    DAOP6_NETWORK_NODES.forEach((n, idx) => {
        nodeMap[n.key] = { kelas: n.kelas, line: n.line, routeIndex: idx };
    });

    return Object.keys(STATION_META).map((key, rawIdx) => {
        const meta = STATION_META[key];
        const count = (PARSED_DATABASE[key] || []).length;
        const netInfo = nodeMap[key] || { kelas: 'Kelas III', line: 'Wilayah Daop 6', routeIndex: rawIdx + 100 };
        const cleanName = meta.name.replace(/^Stasiun\s+/i, '');

        return {
            key,
            name: meta.name,
            cleanName,
            code: meta.code,
            count,
            isActive: count > 0,
            kelas: netInfo.kelas,
            line: netInfo.line,
            routeIndex: netInfo.routeIndex,
            classRank: CLASS_RANK_MAP[netInfo.kelas] || 99
        };
    });
}

// 1. CUSTOM SEARCHABLE DROPDOWN DI HALAMAN JADWAL (URUT ABJAD A-Z)
const csdContainer = document.getElementById('custom-station-dropdown');
const csdTrigger = document.getElementById('station-custom-trigger');
const csdSelectedText = document.getElementById('csd-selected-text');
const csdMenu = document.getElementById('csd-menu');
const csdSearchInput = document.getElementById('csd-search-input');
const csdOptionsList = document.getElementById('csd-options-list');

function syncSelectedStationLabel(stationKey) {
    const meta = STATION_META[stationKey];
    if (!meta || !csdSelectedText) return;
    const isActive = (PARSED_DATABASE[stationKey] || []).length > 0;
    csdSelectedText.textContent = `${meta.name} (${meta.code}) — ${isActive ? 'Aktif' : '[Segera Hadir]'}`;
}

function renderCustomStationDropdown(filterText = '') {
    if (!csdOptionsList) return;
    const q = filterText.trim().toLowerCase();

    const sortedStations = getAllStationsList()
        .sort((a, b) => a.cleanName.localeCompare(b.cleanName, 'id'))
        .filter(st =>
            !q ||
            st.name.toLowerCase().includes(q) ||
            st.code.toLowerCase().includes(q)
        );

    if (sortedStations.length === 0) {
        csdOptionsList.innerHTML = `<li class="csd-empty">Stasiun "${filterText}" tidak ditemukan.</li>`;
        return;
    }

    csdOptionsList.innerHTML = sortedStations.map(st => {
        const isSelected = st.key === currentStation;
        return `
            <li class="csd-option ${isSelected ? 'selected' : ''}" data-key="${st.key}">
                <div class="csd-opt-main">
                    <strong>${st.name} <span class="csd-code">(${st.code})</span></strong>
                    <small>${st.kelas} • ${st.isActive ? `${st.count} perjalanan KA` : 'Jadwal belum diinput'}</small>
                </div>
                <span class="status-badge ${st.isActive ? 'status-berhenti' : 'status-langsung'}">
                    ${st.isActive ? 'AKTIF' : 'SOON'}
                </span>
            </li>
        `;
    }).join('');

    csdOptionsList.querySelectorAll('.csd-option').forEach(li => {
        li.addEventListener('click', () => {
            const chosenKey = li.dataset.key;
            saveSelectedStation(chosenKey);
            syncSelectedStationLabel(chosenKey);
            toggleCustomDropdown(false);
            updateFullScheduleView();
        });
    });
}

function toggleCustomDropdown(open) {
    if (!csdMenu || !csdTrigger) return;
    if (open) {
        csdMenu.classList.remove('hidden');
        csdTrigger.setAttribute('aria-expanded', 'true');
        csdSearchInput.value = '';
        renderCustomStationDropdown('');
        setTimeout(() => csdSearchInput.focus(), 30);
    } else {
        csdMenu.classList.add('hidden');
        csdTrigger.setAttribute('aria-expanded', 'false');
    }
}

if (csdTrigger) {
    csdTrigger.addEventListener('click', () => {
        const isOpen = !csdMenu.classList.contains('hidden');
        toggleCustomDropdown(!isOpen);
    });

    csdSearchInput.addEventListener('input', (e) => {
        renderCustomStationDropdown(e.target.value);
    });

    document.addEventListener('click', (e) => {
        if (csdContainer && !csdContainer.contains(e.target)) {
            toggleCustomDropdown(false);
        }
    });
}

// 2. DIREKTORI STASIUN DAOP 6 (PENCARIAN + FILTER STATUS + MULTI-SORT)
let corridorSearchQuery = '';
let corridorStatusFilter = 'ALL';
let corridorSortMode = 'ACTIVE_AZ';

const corridorSearchInput = document.getElementById('corridor-search-input');
const corridorSortSelect = document.getElementById('corridor-sort-select');
const corridorStatusPills = document.querySelectorAll('#corridor-status-pills .pill');

function renderCorridorAndHomeChips() {
    const container = document.getElementById('corridor-list');
    const homeChipsContainer = document.getElementById('home-active-stations');
    const allStations = getAllStationsList();

    const activeCount = allStations.filter(s => s.isActive).length;
    const soonCount = allStations.length - activeCount;
    corridorStatusPills.forEach(pill => {
        if (pill.dataset.status === 'ALL') pill.textContent = `Semua Stasiun (${allStations.length})`;
        if (pill.dataset.status === 'ACTIVE') pill.textContent = `Database Aktif (${activeCount})`;
        if (pill.dataset.status === 'SOON') pill.textContent = `Segera Hadir (${soonCount})`;
    });

    const q = corridorSearchQuery.trim().toLowerCase();
    let filtered = allStations.filter(st => {
        const matchStatus =
            corridorStatusFilter === 'ALL' ||
            (corridorStatusFilter === 'ACTIVE' && st.isActive) ||
            (corridorStatusFilter === 'SOON' && !st.isActive);

        const matchQuery =
            !q ||
            st.name.toLowerCase().includes(q) ||
            st.code.toLowerCase().includes(q) ||
            st.kelas.toLowerCase().includes(q);

        return matchStatus && matchQuery;
    });

    filtered.sort((a, b) => {
        if (corridorSortMode === 'ACTIVE_AZ') {
            if (a.isActive !== b.isActive) return a.isActive ? -1 : 1;
            return a.cleanName.localeCompare(b.cleanName, 'id');
        }
        if (corridorSortMode === 'NAME_AZ') {
            return a.cleanName.localeCompare(b.cleanName, 'id');
        }
        if (corridorSortMode === 'NAME_ZA') {
            return b.cleanName.localeCompare(a.cleanName, 'id');
        }
        if (corridorSortMode === 'CLASS_HIERARCHY') {
            if (a.classRank !== b.classRank) return a.classRank - b.classRank;
            return a.cleanName.localeCompare(b.cleanName, 'id');
        }
        if (corridorSortMode === 'ROUTE_ORDER') {
            return a.routeIndex - b.routeIndex;
        }
        return 0;
    });

    if (container) {
        if (filtered.length === 0) {
            container.innerHTML = `<div class="empty-state" style="grid-column: 1 / -1;"><p>Tidak ada stasiun yang cocok dengan filter pencarian.</p></div>`;
        } else {
            container.innerHTML = filtered.map(st => `
                <div class="corridor-item" onclick="selectStationFromCorridor('${st.key}')">
                    <div>
                        <strong>${st.name} (${st.code})</strong>
                        <div style="font-size:0.78rem; color:#2563eb; font-weight:700; margin-top:2px;">
                            ${st.kelas}
                        </div>
                        <div style="font-size:0.8rem; color:#64748b; margin-top:3px;">
                            ${st.isActive ? `${st.count} perjalanan KA terdaftar` : 'Database menunggu input'}
                        </div>
                    </div>
                    <span class="status-badge ${st.isActive ? 'status-berhenti' : 'status-langsung'}">
                        ${st.isActive ? 'AKTIF' : 'SOON'}
                    </span>
                </div>
            `).join('');
        }
    }

    if (homeChipsContainer) {
        const activeSortedAZ = allStations
            .filter(s => s.isActive)
            .sort((a, b) => a.cleanName.localeCompare(b.cleanName, 'id'));

        homeChipsContainer.innerHTML = activeSortedAZ.map(st => `
            <button class="quick-chip" onclick="selectStationFromCorridor('${st.key}')">
                ${st.cleanName} (${st.code}) &rarr;
            </button>
        `).join('');
    }
}

if (corridorSearchInput) {
    corridorSearchInput.addEventListener('input', (e) => {
        corridorSearchQuery = e.target.value;
        renderCorridorAndHomeChips();
    });
}

if (corridorSortSelect) {
    corridorSortSelect.addEventListener('change', (e) => {
        corridorSortMode = e.target.value;
        renderCorridorAndHomeChips();
    });
}

corridorStatusPills.forEach(pill => {
    pill.addEventListener('click', () => {
        corridorStatusPills.forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        corridorStatusFilter = pill.dataset.status;
        renderCorridorAndHomeChips();
    });
});

window.selectStationFromCorridor = function(stationKey) {
    saveSelectedStation(stationKey);
    currentSearchQuery = '';
    if (searchKaInput) searchKaInput.value = '';
    syncSelectedStationLabel(stationKey);
    renderCustomStationDropdown('');
    switchPage('jadwal');
    updateFullScheduleView();
};

window.openStationFromHomeSearch = function(stationKey, trainFilterQuery) {
    saveSelectedStation(stationKey);
    currentSearchQuery = trainFilterQuery || '';
    if (searchKaInput) searchKaInput.value = currentSearchQuery;
    syncSelectedStationLabel(stationKey);
    renderCustomStationDropdown('');

    if (homeSearchResults) homeSearchResults.classList.add('hidden');
    switchPage('jadwal');
    updateFullScheduleView();
};

/* ==========================================================================
   EVENT LISTENERS FILTER TABEL & JAM REAL-TIME
   ========================================================================== */
if (searchKaInput) {
    searchKaInput.addEventListener('input', (e) => {
        currentSearchQuery = e.target.value;
        updateFullScheduleView();
    });
}

filterPills.forEach(pill => {
    pill.addEventListener('click', () => {
        filterPills.forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        currentStatusFilter = pill.dataset.filter;
        updateFullScheduleView();
    });
});

function formatIndonesianDate(dateObj, upper = false) {
    const hariList = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const bulanList = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

    const hari = hariList[dateObj.getDay()];
    const tanggal = String(dateObj.getDate()).padStart(2, '0');
    const bulan = bulanList[dateObj.getMonth()];
    const tahun = dateObj.getFullYear();

    const formatted = `${hari}, ${tanggal} ${bulan} ${tahun}`;
    return upper ? formatted.toUpperCase() : formatted;
}

function updateClocksAndRealtime() {
    const now = new Date();
    const timeString = [
        String(now.getHours()).padStart(2, '0'),
        String(now.getMinutes()).padStart(2, '0'),
        String(now.getSeconds()).padStart(2, '0')
    ].join(':');

    const dateNormal = formatIndonesianDate(now, false);
    const dateUpper = formatIndonesianDate(now, true);

    const rtClock = document.getElementById('realtime-clock');
    const rtDate = document.getElementById('realtime-date');
    const drClock = document.getElementById('drawer-clock');
    const drDate = document.getElementById('drawer-date');
    const hmClock = document.getElementById('home-clock');
    const hmDate = document.getElementById('home-date');

    if (rtClock) rtClock.textContent = timeString;
    if (rtDate) rtDate.textContent = dateUpper;
    if (drClock) drClock.textContent = `${timeString} WIB`;
    if (drDate) drDate.textContent = dateNormal;
    if (hmClock) hmClock.textContent = timeString;
    if (hmDate) hmDate.textContent = dateNormal;

    renderRealtimeRadar();
}

/* ==========================================================================
   MODUL KRITIK & SARAN (GOOGLE SHEETS API + CADANGAN WHATSAPP)
   ========================================================================== */
const GOOGLE_SHEET_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbx4TFdeRjoyGEQaCvS3K8puRVcC3uAwZG8QPY-PH0uA1WI9Sh7s5pfF01XH3C1qOtg/exec';
const ADMIN_WHATSAPP_NUMBER   = '6283160607606';

const feedbackForm = document.getElementById('feedback-form');
const fbStasiunSelect = document.getElementById('fb-stasiun');
const fbAlertBox = document.getElementById('fb-alert');
const btnSubmitFb = document.getElementById('btn-submit-feedback');
const btnSendWa = document.getElementById('btn-send-wa');

if (fbStasiunSelect) {
    Object.keys(STATION_META)
        .sort((a, b) => STATION_META[a].name.localeCompare(STATION_META[b].name, 'id'))
        .forEach(key => {
            const meta = STATION_META[key];
            const opt = document.createElement('option');
            opt.value = `${meta.name} (${meta.code})`;
            opt.textContent = `${meta.name} (${meta.code})`;
            fbStasiunSelect.appendChild(opt);
        });
}

function showFeedbackAlert(type, message) {
    if (!fbAlertBox) return;
    fbAlertBox.className = `fb-alert ${type}`;
    fbAlertBox.textContent = message;
    fbAlertBox.classList.remove('hidden');
}

if (feedbackForm) {
    feedbackForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const nama = document.getElementById('fb-nama').value.trim() || 'Anonim';
        const kategori = document.getElementById('fb-kategori').value;
        const stasiun = fbStasiunSelect.value;
        const pesan = document.getElementById('fb-pesan').value.trim();
        const honeypot = document.getElementById('fb-honeypot').value;

        if (!pesan) {
            showFeedbackAlert('error', 'Mohon tuliskan isi kritik atau saran terlebih dahulu.');
            return;
        }

        if (!GOOGLE_SHEET_WEBAPP_URL.startsWith('https://script.google.com/')) {
            showFeedbackAlert('error', 'Endpoint Google Sheets belum dikonfigurasi oleh admin. Silakan gunakan tombol Kirim via WhatsApp.');
            return;
        }

        btnSubmitFb.disabled = true;
        btnSubmitFb.textContent = 'Mengirim ke Database...';
        fbAlertBox.classList.add('hidden');

        try {
            await fetch(GOOGLE_SHEET_WEBAPP_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ nama, kategori, stasiun, pesan, website: honeypot })
            });

            showFeedbackAlert('success', 'Terima kasih! Kritik & saran kamu sudah berhasil direkam ke database pengelola.');
            feedbackForm.reset();
        } catch (err) {
            showFeedbackAlert('error', 'Gagal terhubung ke server database. Coba periksa koneksi internet atau kirim via WhatsApp.');
        } finally {
            btnSubmitFb.disabled = false;
            btnSubmitFb.innerHTML = 'Kirim ke Database &rarr;';
        }
    });
}

if (btnSendWa) {
    btnSendWa.addEventListener('click', () => {
        const nama = document.getElementById('fb-nama').value.trim() || 'Anonim';
        const kategori = document.getElementById('fb-kategori').value;
        const stasiun = fbStasiunSelect.value;
        const pesan = document.getElementById('fb-pesan').value.trim();

        if (!pesan) {
            showFeedbackAlert('error', 'Isi dulu kolom detail pesan sebelum mengirim via WhatsApp.');
            return;
        }

        const textWa = [
            `*MASUKAN WEB JADWAL KA DAOP 6*`,
            `• *Nama:* ${nama}`,
            `• *Kategori:* ${kategori}`,
            `• *Stasiun:* ${stasiun}`,
            `• *Pesan:* ${pesan}`
        ].join('\n');

        const waUrl = `https://wa.me/${ADMIN_WHATSAPP_NUMBER}?text=${encodeURIComponent(textWa)}`;
        window.open(waUrl, '_blank', 'noopener,noreferrer');
    });
}

/* ==========================================================================
   RENDER PETA JARINGAN REL DAOP 6 YOGYAKARTA
   ========================================================================== */
function renderDaop6NetworkMap() {
    const wrapper = document.getElementById('network-svg-wrapper');
    if (!wrapper) return;

    const svgNS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(svgNS, 'svg');
    svg.setAttribute('viewBox', '0 0 1580 490');

    const createEl = (tag, attrs) => {
        const el = document.createElementNS(svgNS, tag);
        Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
        return el;
    };

    const dirTexts = [
        { x: 20,   y: 214, text: '← Arah Kutoarjo / Jakarta / Bandung', anchor: 'start', fill: '#60a5fa' },
        { x: 1560, y: 214, text: 'Arah Madiun / Surabaya / Malang →',  anchor: 'end',   fill: '#60a5fa' },
        { x: 1085, y: 20,  text: '↑ Arah Gundih / Semarang / Surabaya Pasarturi', anchor: 'middle', fill: '#fbbf24' }
    ];

    dirTexts.forEach(d => {
        const t = createEl('text', {
            x: d.x, y: d.y, 'text-anchor': d.anchor,
            fill: d.fill, 'font-size': '10.5', 'font-weight': '700',
            'font-family': 'Plus Jakarta Sans, sans-serif'
        });
        t.textContent = d.text;
        svg.appendChild(t);
    });

    const tracks = [
        { d: 'M 35 260 L 1545 260', stroke: '#2563eb', width: 5 },
        { d: 'M 185 260 L 185 355', stroke: '#14b8a6', width: 4.5 },
        { d: 'M 1085 260 L 1085 32', stroke: '#f59e0b', width: 4.5 },
        { d: 'M 1085 195 L 930 195', stroke: '#14b8a6', width: 4.5 },
        { d: 'M 1005 260 L 1165 450', stroke: '#8b5cf6', width: 4.5 }
    ];

    tracks.forEach(tr => {
        svg.appendChild(createEl('path', {
            d: tr.d,
            stroke: tr.stroke,
            'stroke-width': tr.width,
            'stroke-linecap': 'round',
            fill: 'none'
        }));
    });

    DAOP6_NETWORK_NODES.forEach(node => {
        const g = createEl('g', {
            class: `station-node ${node.key === 'YOGYAKARTA' ? 'selected' : ''}`,
            'data-key': node.key
        });

        let r = 4.5, strokeW = 2.2, strokeColor = '#94a3b8', lblClass = 'lbl-kelas3';
        if (node.tier === 'besar') {
            r = 8.5; strokeW = 3.5; strokeColor = '#f97316'; lblClass = 'lbl-besar';
        } else if (node.tier === 'kelas1') {
            r = 6.5; strokeW = 3; strokeColor = '#38bdf8'; lblClass = 'lbl-kelas1';
        } else if (node.tier === 'kelas2') {
            r = 5.5; strokeW = 2.5; strokeColor = '#7dd3fc'; lblClass = 'lbl-kelas2';
        }

        const pulse = createEl('circle', {
            cx: node.x, cy: node.y, r: r + 5,
            fill: 'rgba(249, 115, 22, 0.28)',
            class: 'pulse-ring',
            opacity: node.key === 'YOGYAKARTA' ? '1' : '0'
        });
        g.appendChild(pulse);

        const dot = createEl('circle', {
            cx: node.x, cy: node.y, r: r,
            fill: '#ffffff',
            stroke: strokeColor,
            'stroke-width': strokeW
        });
        g.appendChild(dot);

        let tx = node.x, ty = node.y, anchor = 'middle';
        if (node.lblPos === 'top') {
            ty = node.y - (r + 8);
        } else if (node.lblPos === 'bottom') {
            ty = node.y + (r + 15);
        } else if (node.lblPos === 'right') {
            tx = node.x + (r + 9);
            ty = node.y + 4;
            anchor = 'start';
        }

        const label = createEl('text', {
            x: tx, y: ty,
            'text-anchor': anchor,
            class: lblClass
        });
        label.textContent = node.name;
        g.appendChild(label);

        g.addEventListener('click', () => {
            document.querySelectorAll('.station-node').forEach(el => {
                el.classList.remove('selected');
                const ring = el.querySelector('.pulse-ring');
                if (ring) ring.setAttribute('opacity', '0');
            });
            g.classList.add('selected');
            pulse.setAttribute('opacity', '1');

            updateNetworkDetailCard(node);
        });

        svg.appendChild(g);
    });

    wrapper.innerHTML = '';
    wrapper.appendChild(svg);
}

function updateNetworkDetailCard(node) {
    const badgeEl = document.getElementById('nd-class-badge');
    const nameEl = document.getElementById('nd-station-name');
    const descEl = document.getElementById('nd-station-desc');
    const actionsEl = document.getElementById('nd-actions');

    const hasSchedule = (PARSED_DATABASE[node.key] || []).length > 0;

    if (badgeEl) badgeEl.textContent = node.kelas.toUpperCase();
    if (nameEl) nameEl.textContent = `Stasiun ${node.name} (${node.code})`;
    if (descEl) {
        descEl.textContent = `${node.line} • ${hasSchedule ? `${PARSED_DATABASE[node.key].length} perjalanan KA tercatat dalam rencana GAPEKA 2026.` : 'Database jadwal stasiun ini belum diinput.'}`;
    }

    if (actionsEl) {
        actionsEl.innerHTML = `
            <button class="btn-primary" onclick="selectStationFromCorridor('${node.key}')">
                ${hasSchedule ? 'Buka Jadwal GAPEKA 2026 →' : 'Cek Halaman Jadwal →'}
            </button>
        `;
    }
}

/* ==========================================================================
   PENCARIAN PINTAR DI BERANDA (CARI STASIUN & KERETA API LINTAS DATABASE)
   ========================================================================== */
const homeSearchInput = document.getElementById('home-universal-search');
const homeSearchResults = document.getElementById('home-search-results');
const homeSearchClear = document.getElementById('home-search-clear');

function runHomeUniversalSearch(rawQuery) {
    if (!homeSearchResults || !homeSearchClear) return;
    const q = rawQuery.trim().toLowerCase();
    if (!q) {
        homeSearchResults.classList.add('hidden');
        homeSearchClear.classList.add('hidden');
        return;
    }

    homeSearchClear.classList.remove('hidden');

    const matchedStations = Object.keys(STATION_META)
        .map(key => ({
            key,
            ...STATION_META[key],
            count: (PARSED_DATABASE[key] || []).length
        }))
        .filter(st =>
            st.name.toLowerCase().includes(q) ||
            st.code.toLowerCase().includes(q)
        )
        .slice(0, 4);

    const matchedTrains = [];
    Object.keys(PARSED_DATABASE).forEach(stKey => {
        const list = PARSED_DATABASE[stKey] || [];
        list.forEach(ka => {
            if (
                ka.noKa.toLowerCase().includes(q) ||
                ka.namaKa.toLowerCase().includes(q) ||
                ka.relasi.toLowerCase().includes(q)
            ) {
                matchedTrains.push({
                    stKey,
                    stName: STATION_META[stKey].name,
                    stCode: STATION_META[stKey].code,
                    ...ka
                });
            }
        });
    });

    const topTrains = matchedTrains.slice(0, 6);

    if (matchedStations.length === 0 && topTrains.length === 0) {
        homeSearchResults.innerHTML = `
            <div class="hs-empty">Tidak ditemukan stasiun atau kereta api dengan kata kunci "<strong>${rawQuery}</strong>".</div>
        `;
        homeSearchResults.classList.remove('hidden');
        return;
    }

    let html = '';

    if (matchedStations.length > 0) {
        html += `<div class="hs-group-title">STASIUN WILAYAH DAOP 6</div>`;
        html += matchedStations.map(st => `
            <div class="hs-item" onclick="openStationFromHomeSearch('${st.key}', '')">
                <div>
                    <strong>${st.name} (${st.code})</strong>
                    <small>${st.count > 0 ? `${st.count} perjalanan KA tersedia` : 'Database jadwal belum diinput'}</small>
                </div>
                <span class="hs-badge">${st.count > 0 ? 'Buka Jadwal →' : 'Segera Hadir'}</span>
            </div>
        `).join('');
    }

    if (topTrains.length > 0) {
        html += `<div class="hs-group-title">PERJALANAN KERETA API (${matchedTrains.length} DITEMUKAN)</div>`;
        html += topTrains.map(tr => `
            <div class="hs-item" onclick="openStationFromHomeSearch('${tr.stKey}', '${tr.noKa}')">
                <div>
                    <strong>KA ${tr.noKa} — ${tr.namaKa}</strong>
                    <small>Di ${tr.stName} (${tr.stCode}) • Pukul ${tr.time} WIB (${tr.statusKey})</small>
                </div>
                <span class="hs-badge train">Lihat di ${tr.stCode} →</span>
            </div>
        `).join('');
    }

    homeSearchResults.innerHTML = html;
    homeSearchResults.classList.remove('hidden');
}

if (homeSearchInput) {
    homeSearchInput.addEventListener('input', (e) => runHomeUniversalSearch(e.target.value));
    homeSearchInput.addEventListener('focus', (e) => {
        if (e.target.value.trim()) runHomeUniversalSearch(e.target.value);
    });
    if (homeSearchClear) {
        homeSearchClear.addEventListener('click', () => {
            homeSearchInput.value = '';
            homeSearchResults.classList.add('hidden');
            homeSearchClear.classList.add('hidden');
            homeSearchInput.focus();
        });
    }

    document.addEventListener('click', (e) => {
        if (!e.target.closest('.home-search-box') && homeSearchResults) {
            homeSearchResults.classList.add('hidden');
        }
    });
}

/* ==========================================================================
   INISIALISASI AWAL & PEMULIHAN POSISI HALAMAN SAAT REFRESH
   ========================================================================== */
syncSelectedStationLabel(currentStation);
renderCustomStationDropdown('');
renderCorridorAndHomeChips();
renderDaop6NetworkMap();
updateFullScheduleView();
updateClocksAndRealtime();
setInterval(updateClocksAndRealtime, 1000);

// Pulihkan menu aktif berdasarkan URL Hash (#jadwal, #koridor, #audio, dll.) saat refresh
const initialHashPage = window.location.hash.replace('#', '').trim();
if (initialHashPage && document.getElementById(`section-${initialHashPage}`)) {
    switchPage(initialHashPage, false);
}

window.addEventListener('hashchange', () => {
    const hashPage = window.location.hash.replace('#', '').trim() || 'home';
    switchPage(hashPage, false);
});