Chart.register(ChartDataLabels);
Chart.defaults.color = '#64748b';
Chart.defaults.font.family = "'Poppins', sans-serif";

let currentMes = 'Total';
let currentDestino = 'General';
const TOTAL_AFILIADOS = 129302;
const MESES_ORDEN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre'];
const UMBRAL_ALERTA = 450;
// Meses con órdenes cargadas (ord>0): el promedio consultas/derivación se calcula solo sobre ellos
const MESES_CON_ORD = new Set();

// ─── DATOS EXACTOS (fuente: archivos 26-01 al 26-10.xlsx) ────────────────────
const dataEstricta = [
    // ── ENERO (total=740) Amb=622 Int=115 Tele=5 | LR=687 CBA=53 ──
    { mes:'Enero', region:'Córdoba', inst:'Centro Kuspide',                     amb:2,   int:0,   tele:0, ord:0    },
    { mes:'Enero', region:'Córdoba', inst:'Clinica Central de Ojos',            amb:1,   int:0,   tele:0, ord:0    },
    { mes:'Enero', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina',   amb:62,  int:2,   tele:0, ord:51   },
    { mes:'Enero', region:'Córdoba', inst:'Fecundart',                          amb:5,   int:0,   tele:0, ord:0    },
    { mes:'Enero', region:'Córdoba', inst:'Fund. Bulimia y Anorexia',           amb:1,   int:0,   tele:0, ord:0    },
    { mes:'Enero', region:'Córdoba', inst:'Hospital Privado',                   amb:355, int:88,  tele:4, ord:1185 },
    { mes:'Enero', region:'Córdoba', inst:'Instituto Oulton',                   amb:65,  int:2,   tele:0, ord:76   },
    { mes:'Enero', region:'Córdoba', inst:'Instituto Radioterapia',             amb:7,   int:0,   tele:0, ord:0    },
    { mes:'Enero', region:'Córdoba', inst:'Nascentis',                          amb:7,   int:0,   tele:0, ord:0    },
    { mes:'Enero', region:'Córdoba', inst:'Sanatorio Allende',                  amb:86,  int:18,  tele:0, ord:257  },
    { mes:'Enero', region:'Córdoba', inst:'Sanatorio Mayo',                     amb:8,   int:1,   tele:0, ord:46   },
    { mes:'Enero', region:'Córdoba', inst:'Sanatorio Morra',                    amb:0,   int:1,   tele:0, ord:0    },
    { mes:'Enero', region:'Buenos Aires', inst:'Hospital Garrahan',             amb:1,   int:1,   tele:0, ord:0    },
    { mes:'Enero', region:'Buenos Aires', inst:'Hospital Italiano',             amb:20,  int:1,   tele:1, ord:0    },
    { mes:'Enero', region:'Buenos Aires', inst:'Sanatorio Guemes',              amb:1,   int:0,   tele:0, ord:0    },

    // ── FEBRERO (total=598) Amb=497 Int=100 Tele=1 | LR=546 CBA=52 ──
    { mes:'Febrero', region:'Córdoba', inst:'Centro Kuspide',                   amb:1,   int:0,   tele:0, ord:0   },
    { mes:'Febrero', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina', amb:38,  int:6,   tele:0, ord:66  },
    { mes:'Febrero', region:'Córdoba', inst:'Fecundart',                        amb:1,   int:0,   tele:0, ord:0   },
    { mes:'Febrero', region:'Córdoba', inst:'Hospital Privado',                 amb:244, int:63,  tele:1, ord:962 },
    { mes:'Febrero', region:'Córdoba', inst:'Instituto Oulton',                 amb:60,  int:2,   tele:0, ord:72  },
    { mes:'Febrero', region:'Córdoba', inst:'Instituto Radioterapia',           amb:2,   int:0,   tele:0, ord:0   },
    { mes:'Febrero', region:'Córdoba', inst:'Nascentis',                        amb:0,   int:1,   tele:0, ord:0   },
    { mes:'Febrero', region:'Córdoba', inst:'Sanatorio Allende',                amb:118, int:21,  tele:0, ord:224 },
    { mes:'Febrero', region:'Córdoba', inst:'Sanatorio Mayo',                   amb:10,  int:3,   tele:0, ord:39  },
    { mes:'Febrero', region:'Córdoba', inst:'Sanatorio Morra',                  amb:0,   int:1,   tele:0, ord:0   },
    { mes:'Febrero', region:'Buenos Aires', inst:'Hospital Garrahan',           amb:2,   int:0,   tele:0, ord:0   },
    { mes:'Febrero', region:'Buenos Aires', inst:'Hospital Italiano',           amb:20,  int:4,   tele:0, ord:0   },

    // ── MARZO (total=834) Amb=700 Int=133 Tele=1 | LR=811 CBA=23 ──
    { mes:'Marzo', region:'Córdoba', inst:'Centro Kuspide',                     amb:1,   int:1,   tele:0, ord:0    },
    { mes:'Marzo', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina',   amb:60,  int:3,   tele:0, ord:54   },
    { mes:'Marzo', region:'Córdoba', inst:'Fecundart',                          amb:4,   int:0,   tele:0, ord:0    },
    { mes:'Marzo', region:'Córdoba', inst:'Hospital Privado',                   amb:347, int:76,  tele:1, ord:1147 },
    { mes:'Marzo', region:'Córdoba', inst:'Instituto Oulton',                   amb:71,  int:9,   tele:0, ord:82   },
    { mes:'Marzo', region:'Córdoba', inst:'Instituto Radioterapia',             amb:6,   int:1,   tele:0, ord:0    },
    { mes:'Marzo', region:'Córdoba', inst:'Nascentis',                          amb:2,   int:0,   tele:0, ord:0    },
    { mes:'Marzo', region:'Córdoba', inst:'Sanatorio Allende',                  amb:176, int:39,  tele:0, ord:325  },
    { mes:'Marzo', region:'Córdoba', inst:'Sanatorio Mayo',                     amb:14,  int:0,   tele:0, ord:63   },
    { mes:'Marzo', region:'Buenos Aires', inst:'Hospital Italiano',             amb:18,  int:4,   tele:0, ord:0    },
    { mes:'Marzo', region:'Buenos Aires', inst:'Sanatorio Guemes',              amb:1,   int:0,   tele:0, ord:0    },

    // ── ABRIL (total=925) Amb=726 Int=198 Tele=1 | LR=715 CBA=210 ──
    { mes:'Abril', region:'Córdoba', inst:'Centro Kuspide',                     amb:7,   int:1,   tele:0, ord:0    },
    { mes:'Abril', region:'Córdoba', inst:'Clinica Central de Ojos',            amb:1,   int:0,   tele:0, ord:0    },
    { mes:'Abril', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina',   amb:64,  int:3,   tele:0, ord:46   },
    { mes:'Abril', region:'Córdoba', inst:'Clínica Velez Sarsfield',            amb:1,   int:0,   tele:0, ord:0    },
    { mes:'Abril', region:'Córdoba', inst:'Fecundart',                          amb:2,   int:0,   tele:0, ord:0    },
    { mes:'Abril', region:'Córdoba', inst:'Hospital Privado',                   amb:358, int:109, tele:1, ord:1106 },
    { mes:'Abril', region:'Córdoba', inst:'Instituto Oulton',                   amb:80,  int:7,   tele:0, ord:64   },
    { mes:'Abril', region:'Córdoba', inst:'Instituto Radioterapia',             amb:9,   int:2,   tele:0, ord:0    },
    { mes:'Abril', region:'Córdoba', inst:'Nascentis',                          amb:6,   int:0,   tele:0, ord:0    },
    { mes:'Abril', region:'Córdoba', inst:'Sanatorio Allende',                  amb:169, int:67,  tele:0, ord:340  },
    { mes:'Abril', region:'Córdoba', inst:'Sanatorio Mayo',                     amb:13,  int:2,   tele:0, ord:63   },
    { mes:'Abril', region:'Córdoba', inst:'Sanatorio Morra',                    amb:0,   int:2,   tele:0, ord:0    },
    { mes:'Abril', region:'Buenos Aires', inst:'Hospital Garrahan',             amb:2,   int:1,   tele:0, ord:0    },
    { mes:'Abril', region:'Buenos Aires', inst:'Hospital Italiano',             amb:13,  int:3,   tele:0, ord:0    },
    { mes:'Abril', region:'Buenos Aires', inst:'Sanatorio Guemes',              amb:2,   int:0,   tele:0, ord:0    },

    // ── MAYO (total=453) Amb=338 Int=115 Tele=0 | LR=279 CBA=174 ──
    { mes:'Mayo', region:'Córdoba', inst:'Centro Castillo Morales',             amb:1,   int:0,   tele:0, ord:0   },
    { mes:'Mayo', region:'Córdoba', inst:'Centro Kuspide',                      amb:2,   int:0,   tele:0, ord:0   },
    { mes:'Mayo', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina',    amb:68,  int:7,   tele:0, ord:72  },
    { mes:'Mayo', region:'Córdoba', inst:'Fecundart',                           amb:2,   int:3,   tele:0, ord:0   },
    { mes:'Mayo', region:'Córdoba', inst:'Hospital Privado',                    amb:119, int:69,  tele:0, ord:489 },
    { mes:'Mayo', region:'Córdoba', inst:'Instituto Oulton',                    amb:75,  int:7,   tele:0, ord:74  },
    { mes:'Mayo', region:'Córdoba', inst:'Instituto Radioterapia',              amb:2,   int:2,   tele:0, ord:0   },
    { mes:'Mayo', region:'Córdoba', inst:'Nascentis',                           amb:5,   int:1,   tele:0, ord:0   },
    { mes:'Mayo', region:'Córdoba', inst:'Sanatorio Allende',                   amb:50,  int:18,  tele:0, ord:120 },
    { mes:'Mayo', region:'Córdoba', inst:'Sanatorio Mayo',                      amb:12,  int:3,   tele:0, ord:67  },
    { mes:'Mayo', region:'Córdoba', inst:'Sanatorio Morra',                     amb:0,   int:1,   tele:0, ord:0   },
    { mes:'Mayo', region:'Buenos Aires', inst:'Hospital Italiano',              amb:2,   int:1,   tele:0, ord:0   },
    { mes:'Mayo', region:'Buenos Aires', inst:'Sanatorio Guemes',               amb:0,   int:3,   tele:0, ord:0   },

    // ── JULIO (total=814) Amb=619 Int=195 Tele=0 | LR=632 CBA=182 ──
    { mes:'Julio', region:'Córdoba',      inst:'Centro Kuspide',                   amb:8,   int:0,   tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Clinica de Ojos Reyes-Giobellina', amb:84,  int:3,   tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Fecundart',                        amb:4,   int:0,   tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Hospital Privado',                 amb:351, int:134, tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Instituto Oulton',                 amb:66,  int:11,  tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Instituto Radioterapia',           amb:7,   int:3,   tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Nascentis',                        amb:1,   int:0,   tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Sanatorio Allende',                amb:80,  int:34,  tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Sanatorio Mayo',                   amb:14,  int:5,   tele:0, ord:0 },
    { mes:'Julio', region:'Córdoba',      inst:'Sanatorio Morra',                  amb:1,   int:1,   tele:0, ord:0 },
    { mes:'Julio', region:'Buenos Aires', inst:'Hospital Garrahan',                amb:2,   int:0,   tele:0, ord:0 },
    { mes:'Julio', region:'Buenos Aires', inst:'Hospital Italiano',                amb:0,   int:1,   tele:0, ord:0 },
    { mes:'Julio', region:'Buenos Aires', inst:'Sanatorio Guemes',                 amb:1,   int:3,   tele:0, ord:0 },
    // Órdenes: fuente resumen 26-06.xlsx actualizado
    { mes:'Junio', region:'Córdoba', inst:'Centro Kuspide',                     amb:3,   int:2,   tele:0, ord:0   },
    { mes:'Junio', region:'Córdoba', inst:'Clinica Central de Ojos',            amb:1,   int:0,   tele:0, ord:0   },
    { mes:'Junio', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina',   amb:82,  int:4,   tele:0, ord:58  },
    { mes:'Junio', region:'Córdoba', inst:'Fecundart',                          amb:2,   int:0,   tele:0, ord:0   },
    { mes:'Junio', region:'Córdoba', inst:'Hospital Privado',                   amb:272, int:121, tele:0, ord:805 },
    { mes:'Junio', region:'Córdoba', inst:'Instituto Oulton',                   amb:68,  int:9,   tele:0, ord:68  },
    { mes:'Junio', region:'Córdoba', inst:'Instituto Radioterapia',             amb:7,   int:1,   tele:0, ord:0   },
    { mes:'Junio', region:'Córdoba', inst:'Nascentis',                          amb:2,   int:1,   tele:0, ord:0   },
    { mes:'Junio', region:'Córdoba', inst:'Sanatorio Allende',                  amb:68,  int:29,  tele:0, ord:156 },
    { mes:'Junio', region:'Córdoba', inst:'Sanatorio Mayo',                     amb:11,  int:5,   tele:0, ord:66  },
    { mes:'Junio', region:'Córdoba', inst:'Sanatorio Morra',                    amb:1,   int:0,   tele:0, ord:4   },
    { mes:'Junio', region:'Buenos Aires', inst:'Hospital Italiano',             amb:2,   int:0,   tele:0, ord:0   },
    { mes:'Junio', region:'Buenos Aires', inst:'Sanatorio Guemes',              amb:0,   int:2,   tele:0, ord:0   },

    // ── AGOSTO (total=837) Amb=654 Int=181 Tele=2 | LR=693 CBA=144 | fuente 26-08.xlsx ──
    { mes:'Agosto', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina', amb:96, int:6, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Fecundart', amb:2, int:0, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Fund. Bulimia y Anorexia', amb:1, int:0, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Hospital Privado', amb:365, int:126, tele:2, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Instituto Oulton', amb:75, int:10, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Instituto Radioterapia', amb:10, int:3, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Nascentis', amb:1, int:0, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Sanatorio Allende', amb:92, int:27, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Sanatorio Mayo', amb:7, int:5, tele:0, ord:0 },
    { mes:'Agosto', region:'Córdoba', inst:'Sanatorio Morra', amb:1, int:3, tele:0, ord:0 },
    { mes:'Agosto', region:'Buenos Aires', inst:'Hospital Italiano', amb:3, int:1, tele:0, ord:0 },
    { mes:'Agosto', region:'Buenos Aires', inst:'Sanatorio Guemes', amb:1, int:0, tele:0, ord:0 },

    // ── SEPTIEMBRE (total=933) Amb=742 Int=189 Tele=2 | LR=774 CBA=150 Sin origen=9 | fuente 26-09.xlsx ──
    { mes:'Septiembre', region:'Córdoba', inst:'Centro Kuspide', amb:2, int:0, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Clinica Central de Ojos', amb:2, int:0, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina', amb:105, int:10, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Fecundart', amb:5, int:0, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Hospital Privado', amb:408, int:128, tele:1, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Instituto Oulton', amb:68, int:9, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Instituto Radioterapia', amb:9, int:2, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Nascentis', amb:6, int:0, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Sanatorio Allende', amb:100, int:32, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Sanatorio Mayo', amb:24, int:5, tele:0, ord:0 },
    { mes:'Septiembre', region:'Córdoba', inst:'Sanatorio Morra', amb:0, int:1, tele:0, ord:0 },
    { mes:'Septiembre', region:'Buenos Aires', inst:'Hospital Alemán', amb:1, int:0, tele:0, ord:0 },
    { mes:'Septiembre', region:'Buenos Aires', inst:'Hospital Garrahan', amb:2, int:0, tele:0, ord:0 },
    { mes:'Septiembre', region:'Buenos Aires', inst:'Hospital Italiano', amb:8, int:1, tele:1, ord:0 },
    { mes:'Septiembre', region:'Buenos Aires', inst:'Sanatorio Guemes', amb:2, int:1, tele:0, ord:0 },

    // ── OCTUBRE (total=971) Amb=791 Int=178 Tele=2 | LR=855 CBA=116 | fuente 26-10.xlsx ──
    { mes:'Octubre', region:'Córdoba', inst:'Centro Kuspide', amb:1, int:0, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Clinica Central de Ojos', amb:1, int:0, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Clinica de Ojos Reyes-Giobellina', amb:113, int:4, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Clínica Velez Sarsfield', amb:15, int:7, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Fecundart', amb:2, int:0, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Hospital Privado', amb:434, int:120, tele:1, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Instituto Oulton', amb:102, int:7, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Instituto Radioterapia', amb:9, int:2, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Nascentis', amb:3, int:3, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Sanatorio Allende', amb:81, int:29, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Sanatorio Mayo', amb:17, int:3, tele:0, ord:0 },
    { mes:'Octubre', region:'Córdoba', inst:'Sanatorio Morra', amb:1, int:0, tele:0, ord:0 },
    { mes:'Octubre', region:'Buenos Aires', inst:'Hospital Alemán', amb:3, int:1, tele:0, ord:0 },
    { mes:'Octubre', region:'Buenos Aires', inst:'Hospital Garrahan', amb:1, int:1, tele:0, ord:0 },
    { mes:'Octubre', region:'Buenos Aires', inst:'Hospital Italiano', amb:7, int:0, tele:1, ord:0 },
    { mes:'Octubre', region:'Buenos Aires', inst:'Sanatorio Guemes', amb:1, int:1, tele:0, ord:0 },
];

// ─── ORIGEN: derivaciones por sede (fuente exacta detalle xlsx) ──────────────
const origenData = {
    'Enero':   { lr: 687, cba: 53,  total: 740 },
    'Febrero': { lr: 546, cba: 52,  total: 598 },
    'Marzo':   { lr: 811, cba: 23,  total: 834 },
    'Abril':   { lr: 715, cba: 210, total: 925 },
    'Mayo':    { lr: 279, cba: 174, total: 453 },
    'Julio':   { lr: 632, cba: 182, total: 814 },
    'Agosto': { lr: 693, cba: 144, total: 837 },
    'Septiembre': { lr: 774, cba: 150, sin: 9, total: 933 },
    'Octubre': { lr: 855, cba: 116, total: 971 },
};

// ─── ACTIVIDAD DIARIA POR MES (fuente exacta detalle xlsx) ───────────────────
const actividadDiaria = {
    'Enero': [
        {d:'01/01 Jue',v:4},{d:'02/01 Vie',v:74},{d:'05/01 Lun',v:48},{d:'06/01 Mar',v:44},
        {d:'07/01 Mié',v:47},{d:'08/01 Jue',v:36},{d:'09/01 Vie',v:21},{d:'11/01 Dom',v:1},
        {d:'12/01 Lun',v:43},{d:'13/01 Mar',v:53},{d:'14/01 Mié',v:25},{d:'15/01 Jue',v:27},
        {d:'16/01 Vie',v:43},{d:'19/01 Lun',v:29},{d:'20/01 Mar',v:32},{d:'21/01 Mié',v:29},
        {d:'22/01 Jue',v:19},{d:'23/01 Vie',v:26},{d:'24/01 Sáb',v:3},{d:'25/01 Dom',v:2},
        {d:'26/01 Lun',v:29},{d:'27/01 Mar',v:38},{d:'28/01 Mié',v:18},{d:'29/01 Jue',v:25},
        {d:'30/01 Vie',v:22},{d:'31/01 Sáb',v:2},
    ],
    'Febrero': [
        {d:'01/02 Dom',v:2},{d:'02/02 Lun',v:32},{d:'03/02 Mar',v:19},{d:'04/02 Mié',v:25},
        {d:'05/02 Jue',v:35},{d:'06/02 Vie',v:25},{d:'07/02 Sáb',v:1},{d:'09/02 Lun',v:36},
        {d:'10/02 Mar',v:40},{d:'11/02 Mié',v:31},{d:'12/02 Jue',v:36},{d:'13/02 Vie',v:37},
        {d:'15/02 Dom',v:1},{d:'17/02 Mar',v:1},{d:'18/02 Mié',v:19},{d:'19/02 Jue',v:40},
        {d:'20/02 Vie',v:56},{d:'21/02 Sáb',v:3},{d:'22/02 Dom',v:1},{d:'23/02 Lun',v:47},
        {d:'24/02 Mar',v:17},{d:'25/02 Mié',v:30},{d:'26/02 Jue',v:27},{d:'27/02 Vie',v:36},
        {d:'28/02 Sáb',v:1},
    ],
    'Marzo': [
        {d:'01/03 Dom',v:4},{d:'02/03 Lun',v:35},{d:'03/03 Mar',v:35},{d:'04/03 Mié',v:38},
        {d:'05/03 Jue',v:30},{d:'06/03 Vie',v:31},{d:'09/03 Lun',v:29},{d:'10/03 Mar',v:39},
        {d:'11/03 Mié',v:37},{d:'12/03 Jue',v:35},{d:'13/03 Vie',v:53},{d:'14/03 Sáb',v:1},
        {d:'15/03 Dom',v:2},{d:'16/03 Lun',v:62},{d:'17/03 Mar',v:51},{d:'18/03 Mié',v:37},
        {d:'19/03 Jue',v:37},{d:'20/03 Vie',v:42},{d:'22/03 Dom',v:2},{d:'23/03 Lun',v:2},
        {d:'24/03 Mar',v:2},{d:'25/03 Mié',v:38},{d:'26/03 Jue',v:37},{d:'27/03 Vie',v:71},
        {d:'29/03 Dom',v:1},{d:'30/03 Lun',v:43},{d:'31/03 Mar',v:41},
    ],
    'Abril': [
        {d:'01/04 Mié',v:46},{d:'03/04 Vie',v:4},{d:'06/04 Lun',v:56},{d:'07/04 Mar',v:54},
        {d:'08/04 Mié',v:53},{d:'09/04 Jue',v:41},{d:'10/04 Vie',v:46},{d:'11/04 Sáb',v:1},
        {d:'13/04 Lun',v:47},{d:'14/04 Mar',v:44},{d:'15/04 Mié',v:47},{d:'16/04 Jue',v:32},
        {d:'17/04 Vie',v:41},{d:'20/04 Lun',v:36},{d:'21/04 Mar',v:46},{d:'22/04 Mié',v:48},
        {d:'23/04 Jue',v:55},{d:'24/04 Vie',v:43},{d:'27/04 Lun',v:41},{d:'28/04 Mar',v:43},
        {d:'29/04 Mié',v:55},{d:'30/04 Jue',v:46},
    ],
    'Mayo': [
        {d:'01/05 Vie',v:1},{d:'02/05 Sáb',v:2},{d:'04/05 Lun',v:30},{d:'05/05 Mar',v:26},
        {d:'06/05 Mié',v:11},{d:'07/05 Jue',v:30},{d:'08/05 Vie',v:26},{d:'09/05 Sáb',v:1},
        {d:'11/05 Lun',v:27},{d:'12/05 Mar',v:27},{d:'13/05 Mié',v:23},{d:'14/05 Jue',v:14},
        {d:'15/05 Vie',v:22},{d:'16/05 Sáb',v:2},{d:'18/05 Lun',v:25},{d:'19/05 Mar',v:24},
        {d:'20/05 Mié',v:11},{d:'21/05 Jue',v:18},{d:'22/05 Vie',v:13},{d:'26/05 Mar',v:23},
        {d:'27/05 Mié',v:29},{d:'28/05 Jue',v:27},{d:'29/05 Vie',v:41},
    ],
    'Julio': [
        {d:'01/07 Mié',v:79},
        {d:'02/07 Jue',v:53},
        {d:'03/07 Vie',v:60},
        {d:'06/07 Lun',v:36},
        {d:'07/07 Mar',v:22},
        {d:'08/07 Mié',v:29},
        {d:'09/07 Jue',v:2},
        {d:'11/07 Sáb',v:1},
        {d:'12/07 Dom',v:5},
        {d:'13/07 Lun',v:51},
        {d:'14/07 Mar',v:27},
        {d:'15/07 Mié',v:32},
        {d:'16/07 Jue',v:40},
        {d:'17/07 Vie',v:33},
        {d:'18/07 Sáb',v:7},
        {d:'20/07 Lun',v:46},
        {d:'21/07 Mar',v:55},
        {d:'22/07 Mié',v:58},
        {d:'23/07 Jue',v:58},
        {d:'24/07 Vie',v:21},
        {d:'26/07 Dom',v:1},
        {d:'27/07 Lun',v:20},
        {d:'28/07 Mar',v:18},
        {d:'29/07 Mié',v:20},
        {d:'30/07 Jue',v:21},
        {d:'31/07 Vie',v:19},
    ],
    'Junio': [
        {d:'01/06 Lun',v:38},{d:'02/06 Mar',v:36},{d:'03/06 Mié',v:18},{d:'04/06 Jue',v:39},
        {d:'05/06 Vie',v:49},{d:'08/06 Lun',v:29},{d:'09/06 Mar',v:32},{d:'10/06 Mié',v:30},
        {d:'11/06 Jue',v:50},{d:'12/06 Vie',v:30},{d:'13/06 Sáb',v:1},{d:'15/06 Lun',v:4},
        {d:'16/06 Mar',v:21},{d:'17/06 Mié',v:39},{d:'18/06 Jue',v:43},{d:'19/06 Vie',v:48},
        {d:'20/06 Sáb',v:2},{d:'22/06 Lun',v:37},{d:'23/06 Mar',v:50},{d:'24/06 Mié',v:16},
        {d:'25/06 Jue',v:19},{d:'26/06 Vie',v:7},{d:'29/06 Lun',v:29},{d:'30/06 Mar',v:26},
    ],
    'Agosto': [
        {d:'01/08 Sáb',v:99},{d:'02/08 Dom',v:2},{d:'03/08 Lun',v:80},{d:'04/08 Mar',v:4},
        {d:'05/08 Mié',v:61},{d:'06/08 Jue',v:43},{d:'07/08 Vie',v:51},{d:'10/08 Lun',v:42},
        {d:'11/08 Mar',v:44},{d:'12/08 Mié',v:39},{d:'13/08 Jue',v:47},{d:'14/08 Vie',v:43},
        {d:'15/08 Sáb',v:2},{d:'17/08 Lun',v:5},{d:'18/08 Mar',v:53},{d:'19/08 Mié',v:31},
        {d:'20/08 Jue',v:46},{d:'21/08 Vie',v:24},{d:'22/08 Sáb',v:1},{d:'24/08 Lun',v:33},
        {d:'25/08 Mar',v:27},{d:'26/08 Mié',v:24},{d:'27/08 Jue',v:16},{d:'28/08 Vie',v:19},
        {d:'30/08 Dom',v:1},
    ],
    'Septiembre': [
        {d:'01/09 Mar',v:140},{d:'02/09 Mié',v:63},{d:'03/09 Jue',v:31},{d:'04/09 Vie',v:39},
        {d:'05/09 Sáb',v:4},{d:'07/09 Lun',v:46},{d:'08/09 Mar',v:39},{d:'09/09 Mié',v:39},
        {d:'10/09 Jue',v:37},{d:'11/09 Vie',v:36},{d:'14/09 Lun',v:48},{d:'15/09 Mar',v:39},
        {d:'16/09 Mié',v:28},{d:'17/09 Jue',v:44},{d:'18/09 Vie',v:33},{d:'19/09 Sáb',v:1},
        {d:'20/09 Dom',v:3},{d:'21/09 Lun',v:5},{d:'22/09 Mar',v:44},{d:'23/09 Mié',v:37},
        {d:'24/09 Jue',v:37},{d:'25/09 Vie',v:46},{d:'27/09 Dom',v:2},{d:'28/09 Lun',v:22},
        {d:'29/09 Mar',v:40},{d:'30/09 Mié',v:30},
    ],
    'Octubre': [
        {d:'01/10 Jue',v:65},{d:'02/10 Vie',v:30},{d:'03/10 Sáb',v:4},{d:'04/10 Dom',v:1},
        {d:'05/10 Lun',v:68},{d:'06/10 Mar',v:35},{d:'07/10 Mié',v:40},{d:'08/10 Jue',v:34},
        {d:'09/10 Vie',v:41},{d:'10/10 Sáb',v:8},{d:'11/10 Dom',v:3},{d:'12/10 Lun',v:9},
        {d:'13/10 Mar',v:52},{d:'14/10 Mié',v:20},{d:'15/10 Jue',v:126},{d:'16/10 Vie',v:41},
        {d:'17/10 Sáb',v:2},{d:'18/10 Dom',v:3},{d:'19/10 Lun',v:44},{d:'20/10 Mar',v:49},
        {d:'21/10 Mié',v:37},{d:'22/10 Jue',v:37},{d:'23/10 Vie',v:25},{d:'25/10 Dom',v:3},
        {d:'26/10 Lun',v:32},{d:'27/10 Mar',v:51},{d:'28/10 Mié',v:46},{d:'29/10 Jue',v:19},
        {d:'30/10 Vie',v:44},{d:'31/10 Sáb',v:2},
    ],
};

// ─── DIAGNÓSTICOS EXACTOS (fuente: detalle xlsx) ─────────────────────────────
// IMPORTANTE: "total" = total de DERIVACIONES del mes (no solo filas con dx)
// "Otras" absorbe: diagnósticos no categorizados + filas sin diagnóstico
// → la suma de todas las categorías + otras siempre = total derivaciones del mes
// Abril y Mayo no tienen columna DIAGNÓSTICO → null (se muestran como sin datos)
const diagData = {
    'Enero':   { onco:121, trauma:104, nefro:36,  oftalmo:58,  gineco:46,  gastro:43,  neuro:29,  cardio:18,  total:740  },
    'Febrero': { onco:97,  trauma:78,  nefro:30,  oftalmo:40,  gineco:28,  gastro:29,  neuro:31,  cardio:10,  total:598  },
    'Marzo':   { onco:177, trauma:97,  nefro:21,  oftalmo:52,  gineco:43,  gastro:51,  neuro:31,  cardio:17,  total:834  },
    'Julio':   { onco:182, trauma:75, nefro:45,  oftalmo:62,  gineco:35,  gastro:41,  neuro:46,  cardio:26,  total:814  },
    'Mayo':    null,
    'Junio':   { onco:133, trauma:55,  nefro:39,  oftalmo:49,  gineco:49,  gastro:40,  neuro:26,  cardio:26,  total:693  },
    // Ago–Oct: clasificación por palabras clave sobre la columna DIAGNÓSTICO del Detalle; "Otras" absorbe lo no categorizado y las filas sin dx
    'Agosto': { onco:212, trauma:82, nefro:104, oftalmo:85, gineco:41, gastro:57, neuro:40, cardio:50, total:837 },
    'Septiembre': { onco:195, trauma:116, nefro:93, oftalmo:93, gineco:55, gastro:69, neuro:51, cardio:50, total:933 },
    'Octubre': { onco:204, trauma:121, nefro:96, oftalmo:90, gineco:37, gastro:64, neuro:48, cardio:46, total:971 },
};

dataEstricta.forEach(d => { if (d.ord>0) MESES_CON_ORD.add(d.mes); });
let chartInstObj=null, chartDxObj=null, chartOrdenesObj=null, chartDiarioObj=null;

// ─── SETTERS ─────────────────────────────────────────────────────────────────
function setMes(mes) {
    currentMes = mes;
    const btnTotal = document.getElementById('btn-mes-Total');
    const selectMes = document.getElementById('select-mes');
    if (mes === 'Total') {
        btnTotal.classList.add('tab-btn-active'); btnTotal.classList.remove('tab-btn-inactive');
        selectMes.value = 'Total'; selectMes.classList.remove('border-aposTeal','text-aposTeal');
    } else {
        btnTotal.classList.add('tab-btn-inactive'); btnTotal.classList.remove('tab-btn-active');
        selectMes.value = mes; selectMes.classList.add('border-aposTeal','text-aposTeal');
    }
    procesarDatos();
}

function setDestino(destino) {
    currentDestino = destino;
    document.querySelectorAll("[id^='btn-dest-']").forEach(el => el.classList.replace('tab-btn-active','tab-btn-inactive'));
    document.getElementById(`btn-dest-${destino}`).classList.replace('tab-btn-inactive','tab-btn-active');
    procesarDatos();
}

// ─── PROCESAMIENTO PRINCIPAL ──────────────────────────────────────────────────
function procesarDatos() {
    let datos = dataEstricta;
    if (currentMes !== 'Total') datos = datos.filter(d => d.mes === currentMes);
    if (currentDestino !== 'General') datos = datos.filter(d => d.region === currentDestino);

    let sAmb=0, sInt=0, sTele=0, sOrd=0, sAmbOrd=0;
    let instMap = {};
    datos.forEach(row => {
        sAmb+=row.amb; sInt+=row.int; sTele+=row.tele; sOrd+=row.ord;
        const conOrd = MESES_CON_ORD.has(row.mes); if (conOrd) sAmbOrd+=row.amb;
        if (!instMap[row.inst]) instMap[row.inst]={amb:0,int:0,ord:0,total:0,ambOrd:0};
        instMap[row.inst].amb+=row.amb; instMap[row.inst].int+=row.int;
        instMap[row.inst].ord+=row.ord; if (conOrd) instMap[row.inst].ambOrd+=row.amb; instMap[row.inst].total+=row.amb+row.int+row.tele;
    });
    const total = sAmb+sInt+sTele;
    const ambPct  = total>0 ? sAmb/total*100 : 0;
    const intPct  = total>0 ? sInt/total*100 : 0;
    const telePct = total>0 ? sTele/total*100 : 0;

    // ── KPI TOTAL ─────────────────────────────────────────────────────────────
    document.getElementById('kpi-total').innerText = total;
    const pctCont = document.getElementById('kpi-total-pct-container');
    if (currentMes==='Total') { pctCont.style.display='none'; }
    else { pctCont.style.display='flex'; document.getElementById('kpi-total-pct').innerText=`(${(total/TOTAL_AFILIADOS*100).toFixed(2)}%)`; }

    // ── TENDENCIA ─────────────────────────────────────────────────────────────
    const trendRow=document.getElementById('kpi-total-trend-row');
    const trendEl =document.getElementById('kpi-total-trend');
    if (currentMes!=='Total') {
        trendRow.style.display='flex';
        const mesPrev=MESES_ORDEN[MESES_ORDEN.indexOf(currentMes)-1];
        if (mesPrev && dataEstricta.some(d=>d.mes===mesPrev)) {
            let dp=dataEstricta.filter(d=>d.mes===mesPrev);
            if (currentDestino!=='General') dp=dp.filter(d=>d.region===currentDestino);
            const prev=dp.reduce((s,d)=>s+d.amb+d.int+d.tele,0);
            const diff=total-prev, pct=prev>0?((diff/prev)*100).toFixed(1):'—';
            if (diff>0)      { trendEl.className='kpi-trend kpi-trend-up';      trendEl.innerHTML=`<i class="fas fa-arrow-trend-up"></i> +${diff} (${pct}%) vs ${mesPrev}`; }
            else if (diff<0) { trendEl.className='kpi-trend kpi-trend-down';    trendEl.innerHTML=`<i class="fas fa-arrow-trend-down"></i> ${diff} (${pct}%) vs ${mesPrev}`; }
            else             { trendEl.className='kpi-trend kpi-trend-neutral'; trendEl.innerHTML=`<i class="fas fa-minus"></i> Sin cambio`; }
        } else { trendEl.className='kpi-trend kpi-trend-neutral'; trendEl.innerHTML=`<i class="fas fa-minus"></i> Primer mes`; }
    } else { trendRow.style.display='none'; }

    // ── BADGE DATOS PARCIALES ─────────────────────────────────────────────────
    const parcialBadge = document.getElementById('kpi-parcial-badge');
    if (parcialBadge) parcialBadge.style.display = 'none';

    // ── BADGE VOLUMEN ─────────────────────────────────────────────────────────
    const volEl=document.getElementById('kpi-total-vol-alert');
    if (currentMes==='Total'&&currentDestino==='General') {
        const p=total/5400*100;
        if (p>95)       { volEl.style.display='flex'; volEl.className='kpi-balance-alert kpi-balance-warn mt-2'; volEl.innerHTML=`<i class="fas fa-circle-exclamation"></i> ${p.toFixed(1)}% objetivo anual`; }
        else if (p>=81) { volEl.style.display='flex'; volEl.className='kpi-balance-alert kpi-balance-mid mt-2';  volEl.innerHTML=`<i class="fas fa-triangle-exclamation"></i> ${p.toFixed(1)}% objetivo anual`; }
        else            { volEl.style.display='flex'; volEl.className='kpi-balance-alert kpi-balance-ok mt-2';   volEl.innerHTML=`<i class="fas fa-check-circle"></i> ${p.toFixed(1)}% objetivo anual`; }
    } else if (total>450)      { volEl.style.display='flex'; volEl.className='kpi-balance-alert kpi-balance-warn mt-2'; volEl.innerHTML=`<i class="fas fa-circle-exclamation"></i> Alto volumen`; }
    else if (total>=401)       { volEl.style.display='flex'; volEl.className='kpi-balance-alert kpi-balance-mid mt-2';  volEl.innerHTML=`<i class="fas fa-triangle-exclamation"></i> Volumen en alerta`; }
    else                       { volEl.style.display='none'; }

    // ── KPI AMB / INT / TELE ─────────────────────────────────────────────────
    document.getElementById('kpi-amb').innerText  = sAmb;
    document.getElementById('kpi-int').innerText  = sInt;
    document.getElementById('kpi-tele').innerText = sTele;
    document.getElementById('kpi-amb-pct').innerText  = `(${ambPct.toFixed(1)}%)`;
    document.getElementById('kpi-int-pct').innerText  = `(${intPct.toFixed(1)}%)`;
    document.getElementById('kpi-tele-pct').innerText = `(${telePct.toFixed(1)}%)`;

    // ── ALERTAS AMB/INT ───────────────────────────────────────────────────────
    // Referencia: internación ≥20% → rojo | 18–19.9% → amarillo | <18% → verde
    // El ambulatorio es consecuencia del internado (a mayor int% → menor amb%)
    const ambAl=document.getElementById('kpi-amb-balance-alert');
    const intAl=document.getElementById('kpi-int-balance-alert');
    if (total>0) {
        // ── Semáforo INTERNACIÓN (indicador principal) ──
        intAl.style.display='flex';
        if (intPct>=20) {
            intAl.className='kpi-balance-alert kpi-balance-warn';
            intAl.innerHTML=`<i class="fas fa-circle-exclamation"></i> Internación ${intPct.toFixed(1)}% — Excede referencia (≥20%)`;
        } else if (intPct>=18) {
            intAl.className='kpi-balance-alert kpi-balance-mid';
            intAl.innerHTML=`<i class="fas fa-triangle-exclamation"></i> Internación ${intPct.toFixed(1)}% — Próximo al límite (18–19.9%)`;
        } else {
            intAl.className='kpi-balance-alert kpi-balance-ok';
            intAl.innerHTML=`<i class="fas fa-check-circle"></i> Internación ${intPct.toFixed(1)}% — Dentro del rango (<18%)`;
        }
        // ── Semáforo AMBULATORIO (derivado del internado) ──
        ambAl.style.display='flex';
        if (intPct>=20) {
            ambAl.className='kpi-balance-alert kpi-balance-warn';
            ambAl.innerHTML=`<i class="fas fa-arrow-trend-down"></i> Ambulatorio ${ambPct.toFixed(1)}% — Afectado por alta internación`;
        } else if (intPct>=18) {
            ambAl.className='kpi-balance-alert kpi-balance-mid';
            ambAl.innerHTML=`<i class="fas fa-minus"></i> Ambulatorio ${ambPct.toFixed(1)}% — Leve presión por internación`;
        } else {
            ambAl.className='kpi-balance-alert kpi-balance-ok';
            ambAl.innerHTML=`<i class="fas fa-check-circle"></i> Ambulatorio ${ambPct.toFixed(1)}% — Proporción adecuada`;
        }
    } else { ambAl.style.display='none'; intAl.style.display='none'; }

    // ── TENDENCIA AMB/INT ─────────────────────────────────────────────────────
    const ambTr=document.getElementById('kpi-amb-trend');
    const intTr=document.getElementById('kpi-int-trend');
    if (currentMes!=='Total') {
        document.getElementById('kpi-amb-trend-row').style.display='flex';
        document.getElementById('kpi-int-trend-row').style.display='flex';
        const mesPrev=MESES_ORDEN[MESES_ORDEN.indexOf(currentMes)-1];
        if (mesPrev&&dataEstricta.some(d=>d.mes===mesPrev)) {
            let dp=dataEstricta.filter(d=>d.mes===mesPrev);
            if (currentDestino!=='General') dp=dp.filter(d=>d.region===currentDestino);
            const ap=dp.reduce((s,d)=>s+d.amb,0), ip=dp.reduce((s,d)=>s+d.int,0);
            const da=sAmb-ap, pa=ap>0?((da/ap)*100).toFixed(1):'—';
            if (da>0)      { ambTr.className='kpi-trend kpi-trend-up';      ambTr.innerHTML=`<i class="fas fa-arrow-trend-up"></i> +${da} (${pa}%) vs ${mesPrev}`; }
            else if (da<0) { ambTr.className='kpi-trend kpi-trend-down';    ambTr.innerHTML=`<i class="fas fa-arrow-trend-down"></i> ${da} (${pa}%) vs ${mesPrev}`; }
            else           { ambTr.className='kpi-trend kpi-trend-neutral'; ambTr.innerHTML=`<i class="fas fa-minus"></i> Sin cambio`; }
            const di=sInt-ip, pi=ip>0?((di/ip)*100).toFixed(1):'—';
            if (di>0)      { intTr.className='kpi-trend kpi-trend-up';      intTr.innerHTML=`<i class="fas fa-arrow-trend-up"></i> +${di} (${pi}%) vs ${mesPrev}`; }
            else if (di<0) { intTr.className='kpi-trend kpi-trend-down';    intTr.innerHTML=`<i class="fas fa-arrow-trend-down"></i> ${di} (${pi}%) vs ${mesPrev}`; }
            else           { intTr.className='kpi-trend kpi-trend-neutral'; intTr.innerHTML=`<i class="fas fa-minus"></i> Sin cambio`; }
        }
    } else {
        document.getElementById('kpi-amb-trend-row').style.display='none';
        document.getElementById('kpi-int-trend-row').style.display='none';
    }

    // ── SEMÁFOROS ─────────────────────────────────────────────────────────────
    const ktEl=document.getElementById('kpi-total'); const ktCard=document.getElementById('kpi-total-card');
    let tc='';
    if (currentMes==='Total'&&currentDestino==='General') { const p=total/5400*100; tc=p>95?'#ef4444':p>=81?'#f59e0b':'#22c55e'; }
    else { tc=total>450?'#ef4444':total>=401?'#f59e0b':total>=1?'#22c55e':''; }
    ktEl.style.color=tc; ktCard.style.borderLeftColor=tc||'';
    const kiEl=document.getElementById('kpi-int'); const kiCard=document.getElementById('kpi-int-card');
    if (total>0) {
        const intColor = intPct>=20 ? '#ef4444' : intPct>=18 ? '#f59e0b' : '#22c55e';
        kiEl.style.color=intColor; kiCard.style.borderLeftColor=intColor;
    }

    // ── KPI ÓRDENES ───────────────────────────────────────────────────────────
    document.getElementById('kpi-ordenes-total').innerText=sOrd;
    document.getElementById('kpi-ordenes-prom').innerText=sAmbOrd>0?(sOrd/sAmbOrd).toFixed(2):'—';

    // ── ORIGEN ────────────────────────────────────────────────────────────────
    renderOrigen(total);

    // ── ACTIVIDAD DIARIA ──────────────────────────────────────────────────────
    renderActividadDiaria();

    // ── GRÁFICOS ──────────────────────────────────────────────────────────────
    const labels=Object.keys(instMap).sort((a,b)=>instMap[b].total-instMap[a].total);
    renderChartInst(labels, labels.map(l=>instMap[l].amb), labels.map(l=>instMap[l].int));
    const hasOrd=labels.some(l=>instMap[l].ord>0);
    renderChartOrdenes(labels, labels.map(l=>instMap[l].ambOrd>0?(instMap[l].ord/instMap[l].ambOrd).toFixed(2):0), hasOrd);
    renderChartDx();
    renderMonthlyCards();
}

// ─── RENDER ORIGEN ────────────────────────────────────────────────────────────
function renderOrigen(totalGlobal) {
    const el = document.getElementById('origen-panel');
    if (!el) return;

    let lrTotal=0, cbaTotal=0, grandTotal=0, sinTotal=0;
    if (currentMes==='Total') {
        MESES_ORDEN.forEach(m => {
            const o=origenData[m];
            if (o) { lrTotal+=o.lr; cbaTotal+=o.cba; sinTotal+=(o.sin||0); grandTotal+=o.total; }
        });
    } else {
        const o=origenData[currentMes];
        if (o) { lrTotal=o.lr; cbaTotal=o.cba; sinTotal=(o.sin||0); grandTotal=o.total; }
    }

    const lrPct  = grandTotal>0?(lrTotal/grandTotal*100).toFixed(1):'0.0';
    const cbaPct = grandTotal>0?(cbaTotal/grandTotal*100).toFixed(1):'0.0';

    el.innerHTML = `
        <div class="flex items-center gap-2 mb-2">
            <i class="fas fa-map-marker-alt text-aposTeal text-sm"></i>
            <span class="text-xs font-bold uppercase tracking-wider text-slate-500">Origen de Derivaciones</span>
        </div>
        <div class="grid grid-cols-2 gap-3">
            <div class="bg-aposTeal/5 border border-aposTeal/20 rounded-xl p-3 text-center">
                <div class="text-[10px] font-bold uppercase text-slate-400 mb-1"><i class="fas fa-hospital-alt mr-1"></i>Casa Central — La Rioja</div>
                <div class="text-2xl font-extrabold text-aposTeal">${lrTotal.toLocaleString()}</div>
                <div class="text-sm font-bold text-slate-500">${lrPct}%</div>
            </div>
            <div class="bg-aposOrange/5 border border-aposOrange/20 rounded-xl p-3 text-center">
                <div class="text-[10px] font-bold uppercase text-slate-400 mb-1"><i class="fas fa-building mr-1"></i>Delegación — Córdoba</div>
                <div class="text-2xl font-extrabold text-aposOrange">${cbaTotal.toLocaleString()}</div>
                <div class="text-sm font-bold text-slate-500">${cbaPct}%</div>
            </div>
        </div>
        <div class="mt-2 h-2 rounded-full overflow-hidden bg-slate-100 flex">
            <div style="width:${lrPct}%;background:#00a896;" class="transition-all duration-500"></div>
            <div style="width:${cbaPct}%;background:#f59e0b;" class="transition-all duration-500"></div>
        </div>
        <div class="flex justify-between text-[10px] text-slate-400 mt-1">
            <span>La Rioja ${lrPct}%</span><span>Córdoba ${cbaPct}%</span>
        </div>
        ${sinTotal>0?`<div class="text-[10px] text-slate-400 mt-1"><i class="fas fa-circle-info mr-1"></i>${sinTotal} derivaciones sin origen registrado en el detalle</div>`:''}`;
}

// ─── RENDER ACTIVIDAD DIARIA ──────────────────────────────────────────────────
function renderActividadDiaria() {
    const section = document.getElementById('actividad-diaria-section');
    const canvas  = document.getElementById('chartDiario');

    // En Total no es representativo → ocultar toda la sección
    if (currentMes === 'Total') {
        if (section) section.style.display = 'none';
        if (chartDiarioObj) { chartDiarioObj.destroy(); chartDiarioObj = null; }
        return;
    }

    if (section) section.style.display = 'block';

    const dias   = actividadDiaria[currentMes] || [];
    const maxVal = Math.max(...dias.map(d => d.v));

    // Solo el gráfico de barras (sin lista redundante)
    if (canvas) {
        if (chartDiarioObj) chartDiarioObj.destroy();
        chartDiarioObj = new Chart(canvas.getContext('2d'), {
            type: 'bar',
            data: {
                labels: dias.map(d => d.d),
                datasets: [{ label: 'Derivaciones', data: dias.map(d => d.v),
                    backgroundColor: dias.map(d =>
                        d.v === maxVal ? '#f59e0b' :
                        (d.d.includes('Sáb') || d.d.includes('Dom')) ? '#e2e8f0' : '#00a896'),
                    borderRadius: 4 }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                layout: { padding: { top: 22 } },
                plugins: {
                    legend: { display: false },
                    datalabels: { anchor:'end', align:'top', offset:1, color:'#334155',
                        font:{ weight:'700', size:9 }, formatter: v => v > 0 ? v : '', clip: false }
                },
                scales: {
                    x: { ticks: { maxRotation:60, minRotation:60, font:{ size:8 } }, grid:{ display:false } },
                    y: { beginAtZero: true }
                }
            }
        });
    }
}

// ─── CHART INSTITUCIONES ──────────────────────────────────────────────────────
function renderChartInst(labels,amb,int) {
    const ctx=document.getElementById('chartInstituciones').getContext('2d');
    if (chartInstObj) chartInstObj.destroy();
    chartInstObj=new Chart(ctx,{
        type:'bar',
        data:{ labels, datasets:[
            { label:'Ambulatorias', data:amb, backgroundColor:'#00a896', borderRadius:4 },
            { label:'Internación',  data:int, backgroundColor:'#f59e0b', borderRadius:4 }
        ]},
        options:{ responsive:true, maintainAspectRatio:false, layout:{padding:{top:25}},
            scales:{ x:{ticks:{maxRotation:45,minRotation:45,font:{size:10}},grid:{display:false}}, y:{beginAtZero:true} },
            plugins:{ legend:{position:'top',labels:{usePointStyle:true,boxWidth:8,padding:20}},
                datalabels:{anchor:'end',align:'top',offset:2,color:'#334155',font:{weight:'700',size:10},display:ctx=>ctx.dataset.data[ctx.dataIndex]>0,formatter:Math.round,clip:false} } }
    });
}

// ─── CHART ÓRDENES ────────────────────────────────────────────────────────────
function renderChartOrdenes(labels,data,hasData) {
    const ctx=document.getElementById('chartOrdenes').getContext('2d');
    if (chartOrdenesObj) chartOrdenesObj.destroy();
    if (!hasData) {
        chartOrdenesObj=new Chart(ctx,{ type:'bar', data:{labels:[],datasets:[{data:[]}]},
            options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false},datalabels:{display:false},
                title:{display:true,text:'Sin datos de consultas para este período',color:'#94a3b8',font:{size:12,weight:'600'},padding:{top:60}}} } });
        return;
    }
    const fl=labels.filter((_,i)=>parseFloat(data[i])>0);
    const fd=data.filter(v=>parseFloat(v)>0);
    chartOrdenesObj=new Chart(ctx,{ type:'bar',
        data:{ labels:fl, datasets:[{ label:'Consultas/derivación', data:fd, backgroundColor:'#00a896', borderRadius:5 }] },
        options:{ responsive:true, maintainAspectRatio:false, layout:{padding:{top:30}},
            scales:{ x:{ticks:{maxRotation:45,minRotation:45,font:{size:10}}} },
            plugins:{ legend:{display:false}, datalabels:{anchor:'end',align:'top',offset:2,color:'#334155',font:{weight:'700',size:10},formatter:v=>v>0?v:'',clip:false} } }
    });
}

// ─── CHART DIAGNÓSTICOS ───────────────────────────────────────────────────────
function renderChartDx() {
    const ctx = document.getElementById('chartDiagnosticos').getContext('2d');
    if (chartDxObj) chartDxObj.destroy();

    // Acumular meses con datos
    let dd = { onco:0, trauma:0, nefro:0, oftalmo:0, gineco:0, gastro:0, neuro:0, cardio:0, total:0 };
    const meses = currentMes === 'Total' ? MESES_ORDEN : [currentMes];
    let hasDx = false;
    meses.forEach(m => {
        const d = diagData[m];
        if (d) {
            hasDx = true;
            dd.onco   += d.onco;   dd.trauma  += d.trauma;
            dd.nefro  += d.nefro;  dd.oftalmo += d.oftalmo;
            dd.gineco += d.gineco; dd.gastro  += d.gastro;
            dd.neuro  += d.neuro;  dd.cardio  += d.cardio;
            dd.total  += d.total;
        }
    });

    // Otras = total mes - suma categorías (absorbe sin dx + no categorizados)
    const sumCat = dd.onco + dd.trauma + dd.nefro + dd.oftalmo +
                   dd.gineco + dd.gastro + dd.neuro + dd.cardio;
    const otros  = dd.total - sumCat;

    // En vista Total: aclarar que Abril y Mayo no tienen datos de diagnóstico
    const mesesSinDx = currentMes === 'Total'
        ? MESES_ORDEN.filter(m => !diagData[m] && m !== 'Julio')
        : [];
    const totalSinDx = mesesSinDx.reduce((s, m) => {
        const rows = dataEstricta.filter(d => d.mes === m);
        return s + rows.reduce((r, d) => r + d.amb + d.int + d.tele, 0);
    }, 0);

    const tagLabel = hasDx
        ? `Total: ${dd.total.toLocaleString()}${mesesSinDx.length ? ` (de ${(dd.total + totalSinDx).toLocaleString()})` : ''}`
        : 'Sin datos Dx';
    document.getElementById('tag-total-dx').innerText = tagLabel;

    const dxTableEl = document.getElementById('dx-table');
    if (dxTableEl) {
        if (!hasDx) {
            dxTableEl.innerHTML = '<p class="text-xs text-slate-400 py-2">Sin datos de diagnóstico para este mes (Abril / Mayo)</p>';
        } else {
            const rows = [
                { l:'Oncología / Tumores',              v: dd.onco   },
                { l:'Traumatología / Ortopedia',        v: dd.trauma  },
                { l:'Oftalmología',                     v: dd.oftalmo },
                { l:'Ginecología / Mama',               v: dd.gineco  },
                { l:'Gastroenterología / Hepatología',  v: dd.gastro  },
                { l:'Nefrología / Trasplante',          v: dd.nefro   },
                { l:'Neurología / Neurocirugía',        v: dd.neuro   },
                { l:'Cardiología / Vascular',           v: dd.cardio  },
                { l:'Otras patologías',                 v: otros       },
            ].sort((a, b) => b.v - a.v);

            // Verificación: suma debe igualar total
            const sumaVerif = rows.reduce((s, r) => s + r.v, 0);
            const ok = sumaVerif === dd.total;

            dxTableEl.innerHTML = rows.map(r => `
                <div class="flex items-center justify-between text-xs py-0.5 border-b border-slate-50">
                    <span class="text-slate-500">${r.l}</span>
                    <span class="font-bold text-slate-700">${r.v.toLocaleString()}
                        <span class="text-slate-400 font-normal">(${dd.total > 0 ? (r.v / dd.total * 100).toFixed(1) : '0.0'}%)</span>
                    </span>
                </div>`).join('') +
                `<div class="flex items-center justify-between text-xs pt-1 mt-1 border-t-2 border-slate-200 font-bold">
                    <span class="text-slate-700">Total analizado</span>
                    <span class="text-slate-900">${dd.total.toLocaleString()} <span class="text-slate-400 font-normal">(100%)</span></span>
                </div>` +
                (mesesSinDx.length ? `
                <div class="mt-2 text-[10px] text-slate-400 bg-slate-50 rounded p-1.5 leading-tight">
                    <i class="fas fa-circle-info mr-1"></i>
                    <strong>${mesesSinDx.join(' y ')}</strong> (${totalSinDx.toLocaleString()} derivaciones) no cuentan con datos de diagnóstico en el detalle.
                    Total 2026 completo: <strong>${(dd.total + totalSinDx).toLocaleString()}</strong> derivaciones.
                </div>` : '');
        }
    }

    if (!hasDx) {
        chartDxObj = new Chart(ctx, {
            type: 'doughnut',
            data: { labels: ['Sin datos'], datasets: [{ data: [1], backgroundColor: ['#e2e8f0'], borderWidth: 0 }] },
            options: { responsive: true, maintainAspectRatio: false, cutout: '60%',
                plugins: { legend: { display: false }, datalabels: { display: false } } }
        });
        return;
    }

    chartDxObj = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Oncología','Traumatología','Oftalmología','Ginecología/Mama',
                     'Gastroenterología','Nefrología','Neurología','Cardiología','Otras'],
            datasets: [{
                data: [dd.onco, dd.trauma, dd.oftalmo, dd.gineco,
                       dd.gastro, dd.nefro, dd.neuro, dd.cardio, otros],
                backgroundColor: ['#00a896','#f59e0b','#94a3b8','#e07bb5',
                                  '#7c6d8e','#5b9bd5','#007A72','#d95f5f','#cbd5e1'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false, cutout: '60%',
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 9 } } },
                datalabels: {
                    color: '#fff',
                    font: { weight: '800', size: 10 },
                    formatter: v => v > 0 ? v.toLocaleString() : ''
                }
            }
        }
    });
}

// ─── TARJETAS MENSUALES ───────────────────────────────────────────────────────
function renderMonthlyCards() {
    const section=document.getElementById('monthly-cards-section');
    const grid=document.getElementById('monthly-cards-grid');
    if (currentMes!=='Total') { section.style.display='none'; return; }
    section.style.display='block';

    const items=MESES_ORDEN.map(mes=>{
        let d=dataEstricta.filter(x=>x.mes===mes);
        if (currentDestino!=='General') d=d.filter(x=>x.region===currentDestino);
        const amb=d.reduce((s,x)=>s+x.amb,0), int=d.reduce((s,x)=>s+x.int,0), tele=d.reduce((s,x)=>s+x.tele,0);
        const total=amb+int+tele;
        const intPct=total>0?int/total*100:0, ambPct=total>0?amb/total*100:0;
        const o=currentDestino==='General'?origenData[mes]:null;
        return { mes,total,amb,int,tele,intPct,ambPct,o };
    });

    function ic(p) { return p>=20?'#ef4444':p>=18?'#f59e0b':'#22c55e'; }
    function badge(d,p) {
        if (d>0) return `<span class="mc-trend mc-trend-up"><i class="fas fa-arrow-trend-up"></i> +${d} (${p}%)</span>`;
        if (d<0) return `<span class="mc-trend mc-trend-down"><i class="fas fa-arrow-trend-down"></i> ${d} (${p}%)</span>`;
        return `<span class="mc-trend mc-trend-neutral"><i class="fas fa-minus"></i> Sin cambio</span>`;
    }

    grid.innerHTML=items.map((item,idx)=>{
        const prev=idx>0?items[idx-1]:null;
        const isAlert=item.total>=UMBRAL_ALERTA;
        const color=ic(item.intPct);
        let tTrend=`<span class="mc-trend mc-trend-neutral"><i class="fas fa-minus"></i> Primer mes</span>`;
        let aTrend='', iTrend='';
        if (prev) {
            const d=item.total-prev.total, p=prev.total>0?((d/prev.total)*100).toFixed(1):'—'; tTrend=badge(d,p);
            const da=item.amb-prev.amb, pa=prev.amb>0?((da/prev.amb)*100).toFixed(1):'—'; aTrend=badge(da,pa);
            const di=item.int-prev.int, pi=prev.int>0?((di/prev.int)*100).toFixed(1):'—'; iTrend=badge(di,pi);
        }
        const origenHTML=item.o?`
            <div class="mc-divider"></div>
            <div class="mc-sub-row"><span class="mc-sub-label" style="color:#00a896;min-width:3.5rem;font-size:0.6rem;">LR</span>
                <span class="mc-sub-val" style="font-size:0.75rem;">${item.o.lr} <span style="font-size:0.55rem;color:#94a3b8;">(${(item.o.lr/item.o.total*100).toFixed(1)}%)</span></span></div>
            <div class="mc-sub-row"><span class="mc-sub-label" style="color:#f59e0b;min-width:3.5rem;font-size:0.6rem;">CBA</span>
                <span class="mc-sub-val" style="font-size:0.75rem;">${item.o.cba} <span style="font-size:0.55rem;color:#94a3b8;">(${(item.o.cba/item.o.total*100).toFixed(1)}%)</span></span></div>`:''
        ;
        return `
        <div class="monthly-card ${isAlert?'alert':''}" style="border-top-color:${isAlert?'#ef4444':'#00a896'};">
            <div class="month-label">${item.mes}</div>
            <div class="month-total" style="color:${isAlert?'#ef4444':'#1e293b'};">${item.total}</div>
            <div class="month-indicator">${tTrend}</div>
            ${isAlert?`<div class="alert-badge"><i class="fas fa-circle-exclamation"></i> Alto volumen</div>`:''}
            <div class="mc-divider"></div>
            <div class="mc-sub-row">
                <span class="mc-sub-label"><i class="fas fa-walking" style="color:#00a896"></i> Amb.</span>
                <span class="mc-sub-val">${item.amb} <span style="font-size:0.6rem;color:#94a3b8;">(${item.ambPct.toFixed(1)}%)</span></span>
                ${aTrend?`<span class="mc-sub-trend">${aTrend}</span>`:''}
            </div>
            <div class="mc-sub-row">
                <span class="mc-sub-label"><i class="fas fa-bed" style="color:${color}"></i> Int.</span>
                <span class="mc-sub-val" style="color:${color};font-weight:800;">${item.int} <span style="font-size:0.6rem;font-weight:600;">(${item.intPct.toFixed(1)}%)</span></span>
                ${iTrend?`<span class="mc-sub-trend">${iTrend}</span>`:''}
            </div>
            ${origenHTML}
        </div>`;
    }).join('');
}

window.onload = () => procesarDatos();
