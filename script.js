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
    if (currentMes==='Total'&&currentDestino==='General') { const p=total/5400*100; tc=p>95?'#f08a8a':p>=81?'#f2bf78':'#7dd1a0'; }
    else { tc=total>450?'#f08a8a':total>=401?'#f2bf78':total>=1?'#7dd1a0':''; }
    ktEl.style.color=tc; ktCard.style.borderLeftColor=tc||'';
    const kiEl=document.getElementById('kpi-int'); const kiCard=document.getElementById('kpi-int-card');
    if (total>0) {
        const intColor = intPct>=20 ? '#f08a8a' : intPct>=18 ? '#f2bf78' : '#7dd1a0';
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
            <span class="text-xs font-medium uppercase tracking-wider text-slate-500">Origen de Derivaciones</span>
        </div>
        <div class="grid grid-cols-2 gap-3">
            <div class="bg-aposTeal/5 border border-aposTeal/20 rounded-xl p-3 text-center">
                <div class="text-[10px] font-medium uppercase text-slate-400 mb-1"><i class="fas fa-hospital-alt mr-1"></i>Casa Central — La Rioja</div>
                <div class="text-2xl font-normal text-aposTeal">${lrTotal.toLocaleString()}</div>
                <div class="text-sm font-medium text-slate-500">${lrPct}%</div>
            </div>
            <div class="bg-aposOrange/5 border border-aposOrange/20 rounded-xl p-3 text-center">
                <div class="text-[10px] font-medium uppercase text-slate-400 mb-1"><i class="fas fa-building mr-1"></i>Delegación — Córdoba</div>
                <div class="text-2xl font-normal text-aposOrange">${cbaTotal.toLocaleString()}</div>
                <div class="text-sm font-medium text-slate-500">${cbaPct}%</div>
            </div>
        </div>
        <div class="mt-2 h-2 rounded-full overflow-hidden bg-slate-100 flex">
            <div style="width:${lrPct}%;background:#45b5a9;" class="transition-all duration-500"></div>
            <div style="width:${cbaPct}%;background:#f2bf78;" class="transition-all duration-500"></div>
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
                        d.v === maxVal ? '#f2bf78' :
                        (d.d.includes('Sáb') || d.d.includes('Dom')) ? '#e2e8f0' : '#45b5a9'),
                    borderRadius: 4 }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                layout: { padding: { top: 22 } },
                plugins: {
                    legend: { display: false },
                    datalabels: { anchor:'end', align:'top', offset:1, color:'#6b7280',
                        font:{ weight:'500', size:9 }, formatter: v => v > 0 ? v : '', clip: false }
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
    // Ancho mínimo por institución para que las etiquetas no se superpongan (si no entra, se desplaza horizontalmente)
    const wrap=document.getElementById('inst-chart-wrap');
    if (wrap) wrap.style.minWidth = Math.max(labels.length*58, 0) + 'px';
    if (chartInstObj) chartInstObj.destroy();
    chartInstObj=new Chart(ctx,{
        type:'bar',
        data:{ labels, datasets:[
            { label:'Ambulatorias', data:amb, backgroundColor:'#45b5a9', borderRadius:4, categoryPercentage:0.8, barPercentage:0.9 },
            { label:'Internación',  data:int, backgroundColor:'#f2bf78', borderRadius:4, categoryPercentage:0.8, barPercentage:0.9 }
        ]},
        options:{ responsive:true, maintainAspectRatio:false, layout:{padding:{top:25}},
            scales:{ x:{ticks:{maxRotation:45,minRotation:45,font:{size:10}},grid:{display:false}}, y:{beginAtZero:true} },
            plugins:{ legend:{position:'top',labels:{usePointStyle:true,boxWidth:8,padding:20}},
                datalabels:{anchor:'end',align:'top',offset:1,color:'#6b7280',font:{weight:'400',size:9},display:ctx=>ctx.dataset.data[ctx.dataIndex]>0,formatter:Math.round,clip:false} } }
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
        data:{ labels:fl, datasets:[{ label:'Consultas/derivación', data:fd, backgroundColor:'#45b5a9', borderRadius:5 }] },
        options:{ responsive:true, maintainAspectRatio:false, layout:{padding:{top:30}},
            scales:{ x:{ticks:{maxRotation:45,minRotation:45,font:{size:10}}} },
            plugins:{ legend:{display:false}, datalabels:{anchor:'end',align:'top',offset:2,color:'#6b7280',font:{weight:'500',size:10},formatter:v=>v>0?v:'',clip:false} } }
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
                    <span class="font-medium text-slate-500">${r.v.toLocaleString()}
                        <span class="text-slate-400 font-normal">(${dd.total > 0 ? (r.v / dd.total * 100).toFixed(1) : '0.0'}%)</span>
                    </span>
                </div>`).join('') +
                `<div class="flex items-center justify-between text-xs pt-1 mt-1 border-t-2 border-slate-200 font-medium">
                    <span class="text-slate-500">Total analizado</span>
                    <span class="text-slate-500">${dd.total.toLocaleString()} <span class="text-slate-400 font-normal">(100%)</span></span>
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
                backgroundColor: ['#45b5a9','#f2bf78','#94a3b8','#e8a9cb',
                                  '#ab9fbb','#92bde4','#6fb3ab','#e8a0a0','#cbd5e1'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false, cutout: '60%',
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 9 } } },
                datalabels: {
                    color: '#fff',
                    font: { weight:'500', size: 10 },
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

    function ic(p) { return p>=20?'#f08a8a':p>=18?'#f2bf78':'#7dd1a0'; }
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
            <div class="mc-sub-row"><span class="mc-sub-label" style="color:#45b5a9;min-width:3.5rem;font-size:0.6rem;">LR</span>
                <span class="mc-sub-val" style="font-size:0.75rem;">${item.o.lr} <span style="font-size:0.55rem;color:#94a3b8;">(${(item.o.lr/item.o.total*100).toFixed(1)}%)</span></span></div>
            <div class="mc-sub-row"><span class="mc-sub-label" style="color:#f2bf78;min-width:3.5rem;font-size:0.6rem;">CBA</span>
                <span class="mc-sub-val" style="font-size:0.75rem;">${item.o.cba} <span style="font-size:0.55rem;color:#94a3b8;">(${(item.o.cba/item.o.total*100).toFixed(1)}%)</span></span></div>`:''
        ;
        return `
        <div class="monthly-card ${isAlert?'alert':''}" style="border-top-color:${isAlert?'#f08a8a':'#45b5a9'};">
            <div class="month-label">${item.mes}</div>
            <div class="month-total" style="color:${isAlert?'#f08a8a':'#6b7280'};">${item.total}</div>
            <div class="month-indicator">${tTrend}</div>
            ${isAlert?`<div class="alert-badge"><i class="fas fa-circle-exclamation"></i> Alto volumen</div>`:''}
            <div class="mc-divider"></div>
            <div class="mc-sub-row">
                <span class="mc-sub-label"><i class="fas fa-walking" style="color:#45b5a9"></i> Amb.</span>
                <span class="mc-sub-val">${item.amb} <span style="font-size:0.6rem;color:#94a3b8;">(${item.ambPct.toFixed(1)}%)</span></span>
                ${aTrend?`<span class="mc-sub-trend">${aTrend}</span>`:''}
            </div>
            <div class="mc-sub-row">
                <span class="mc-sub-label"><i class="fas fa-bed" style="color:${color}"></i> Int.</span>
                <span class="mc-sub-val" style="color:${color};font-weight: 400;">${item.int} <span style="font-size:0.6rem;font-weight: 400;">(${item.intPct.toFixed(1)}%)</span></span>
                ${iTrend?`<span class="mc-sub-trend">${iTrend}</span>`:''}
            </div>
            ${origenHTML}
        </div>`;
    }).join('');
}

// ═══════════════════════════════════════════════════════════════════════════
// ANÁLISIS DE COSTOS POR TIPO DE DERIVACIÓN
// Fuente: hoja "Detalle" de 26-08, 26-09 y 26-10.xlsx.
// Para ACTUALIZAR: agregar las filas del mes nuevo a costosCat, costosInst y costosOrig
// (misma estructura). El selector de período, gráficos y textos se arman solos.
//   Tipos (k): T Teleconsulta | A Consulta/interconsulta/control | B Solo estudios |
//     C Control+estudios+tratamiento amb. | C2 Control posquirúrgico | D Cirugía amb. |
//     F Internación clínica/tratamiento | E Internación con cirugía
//   flag: 'B' = internación con cirugía de patología benigna/electiva; 'S' = internación
//     clínica con diagnóstico sintomático/"en estudio"; '' = sin marca
//   costosCat  = [mes, tipo, categoría dx, flag, n]
//   costosInst = [mes, tipo, prestador, n]
//   costosOrig = [mes, tipo, origen (LR/CBA/SIN), n]
// ═══════════════════════════════════════════════════════════════════════════
const COSTOS_GRUPOS = [
    {k:'A',  label:'Consulta / interconsulta / control', color:'#8fd0c8'},
    {k:'B',  label:'Solo estudios',                      color:'#b9d9d5'},
    {k:'C',  label:'Control + estudios + tratamiento amb.', color:'#45b5a9'},
    {k:'C2', label:'Control posquirúrgico',              color:'#92bde4'},
    {k:'D',  label:'Cirugía ambulatoria',                color:'#ab9fbb'},
    {k:'F',  label:'Internación clínica / tratamiento',  color:'#f2bf78'},
    {k:'E',  label:'Internación con cirugía',            color:'#f08a8a'},
    {k:'T',  label:'Teleconsulta',                       color:'#cbd5e1'}
];
const COSTOS_PESOS_DEF = { A:1, B:2, C:4, C2:1.5, D:10, F:15, E:40, T:0.7 };
const costosPesos = Object.assign({}, COSTOS_PESOS_DEF);
const COSTOS_CAT = { onco:'Oncología', trauma:'Traumatología', nefro:'Nefrología', oftalmo:'Oftalmología', gineco:'Ginecología', gastro:'Gastroenterología', neuro:'Neurología', cardio:'Cardiología', otras:'Otras' };
const COSTOS_ORG = { LR:'La Rioja', CBA:'Córdoba', SIN:'Sin origen registrado' };
let costosUnidad = 'indice', costosPeriodo = 'Total';
let chartCostoTipoObj=null, chartCostoMesObj=null;

const costosCat = [
    // Agosto
    ['Agosto','T','gineco','',1],
    ['Agosto','T','onco','',1],
    ['Agosto','A','cardio','',12],
    ['Agosto','A','gastro','',10],
    ['Agosto','A','gineco','',18],
    ['Agosto','A','nefro','',15],
    ['Agosto','A','neuro','',16],
    ['Agosto','A','oftalmo','',16],
    ['Agosto','A','onco','',32],
    ['Agosto','A','otras','',40],
    ['Agosto','A','trauma','',22],
    ['Agosto','B','cardio','',4],
    ['Agosto','B','gastro','',1],
    ['Agosto','B','gineco','',6],
    ['Agosto','B','nefro','',1],
    ['Agosto','B','neuro','',2],
    ['Agosto','B','onco','',21],
    ['Agosto','B','otras','',2],
    ['Agosto','B','trauma','',1],
    ['Agosto','C','cardio','',16],
    ['Agosto','C','gastro','',20],
    ['Agosto','C','gineco','',9],
    ['Agosto','C','nefro','',62],
    ['Agosto','C','neuro','',14],
    ['Agosto','C','oftalmo','',46],
    ['Agosto','C','onco','',100],
    ['Agosto','C','otras','',90],
    ['Agosto','C','trauma','',31],
    ['Agosto','C2','cardio','',1],
    ['Agosto','C2','gastro','',2],
    ['Agosto','C2','nefro','',2],
    ['Agosto','C2','oftalmo','',4],
    ['Agosto','C2','onco','',6],
    ['Agosto','C2','otras','',1],
    ['Agosto','C2','trauma','',11],
    ['Agosto','D','cardio','',1],
    ['Agosto','D','oftalmo','',15],
    ['Agosto','D','onco','',2],
    ['Agosto','D','otras','',2],
    ['Agosto','F','cardio','',4],
    ['Agosto','F','cardio','S',4],
    ['Agosto','F','gastro','',9],
    ['Agosto','F','gastro','S',2],
    ['Agosto','F','gineco','',1],
    ['Agosto','F','nefro','',12],
    ['Agosto','F','nefro','S',1],
    ['Agosto','F','neuro','',3],
    ['Agosto','F','neuro','S',2],
    ['Agosto','F','onco','',14],
    ['Agosto','F','onco','S',1],
    ['Agosto','F','otras','',12],
    ['Agosto','F','otras','S',4],
    ['Agosto','E','cardio','',7],
    ['Agosto','E','cardio','B',1],
    ['Agosto','E','gastro','',8],
    ['Agosto','E','gastro','B',5],
    ['Agosto','E','gineco','',1],
    ['Agosto','E','gineco','B',5],
    ['Agosto','E','nefro','',9],
    ['Agosto','E','nefro','B',2],
    ['Agosto','E','neuro','',3],
    ['Agosto','E','oftalmo','',3],
    ['Agosto','E','oftalmo','B',1],
    ['Agosto','E','onco','',34],
    ['Agosto','E','onco','B',1],
    ['Agosto','E','otras','',11],
    ['Agosto','E','otras','B',4],
    ['Agosto','E','trauma','',8],
    ['Agosto','E','trauma','B',9],
    // Septiembre
    ['Septiembre','T','neuro','',1],
    ['Septiembre','T','onco','',1],
    ['Septiembre','A','cardio','',15],
    ['Septiembre','A','gastro','',14],
    ['Septiembre','A','gineco','',13],
    ['Septiembre','A','nefro','',17],
    ['Septiembre','A','neuro','',13],
    ['Septiembre','A','oftalmo','',25],
    ['Septiembre','A','onco','',54],
    ['Septiembre','A','otras','',67],
    ['Septiembre','A','trauma','',39],
    ['Septiembre','B','cardio','',2],
    ['Septiembre','B','gineco','',6],
    ['Septiembre','B','neuro','',3],
    ['Septiembre','B','oftalmo','',2],
    ['Septiembre','B','onco','',23],
    ['Septiembre','B','otras','',6],
    ['Septiembre','B','trauma','',1],
    ['Septiembre','C','cardio','',15],
    ['Septiembre','C','gastro','',24],
    ['Septiembre','C','gineco','',21],
    ['Septiembre','C','nefro','',53],
    ['Septiembre','C','neuro','',22],
    ['Septiembre','C','oftalmo','',45],
    ['Septiembre','C','onco','',77],
    ['Septiembre','C','otras','',98],
    ['Septiembre','C','trauma','',31],
    ['Septiembre','C2','gastro','',5],
    ['Septiembre','C2','gineco','',3],
    ['Septiembre','C2','nefro','',6],
    ['Septiembre','C2','neuro','',1],
    ['Septiembre','C2','oftalmo','',4],
    ['Septiembre','C2','onco','',4],
    ['Septiembre','C2','otras','',9],
    ['Septiembre','C2','trauma','',17],
    ['Septiembre','D','oftalmo','',7],
    ['Septiembre','F','cardio','',5],
    ['Septiembre','F','cardio','S',3],
    ['Septiembre','F','gastro','',7],
    ['Septiembre','F','gastro','S',3],
    ['Septiembre','F','nefro','',10],
    ['Septiembre','F','neuro','',4],
    ['Septiembre','F','neuro','S',1],
    ['Septiembre','F','oftalmo','',3],
    ['Septiembre','F','onco','',15],
    ['Septiembre','F','otras','',9],
    ['Septiembre','F','otras','S',1],
    ['Septiembre','F','trauma','',5],
    ['Septiembre','E','cardio','',9],
    ['Septiembre','E','cardio','B',1],
    ['Septiembre','E','gastro','',11],
    ['Septiembre','E','gastro','B',5],
    ['Septiembre','E','gineco','',2],
    ['Septiembre','E','gineco','B',10],
    ['Septiembre','E','nefro','',5],
    ['Septiembre','E','nefro','B',2],
    ['Septiembre','E','neuro','',6],
    ['Septiembre','E','oftalmo','',7],
    ['Septiembre','E','onco','',21],
    ['Septiembre','E','otras','',15],
    ['Septiembre','E','otras','B',6],
    ['Septiembre','E','trauma','',16],
    ['Septiembre','E','trauma','B',7],
    // Octubre
    ['Octubre','T','nefro','',1],
    ['Octubre','T','otras','',1],
    ['Octubre','A','cardio','',18],
    ['Octubre','A','gastro','',26],
    ['Octubre','A','gineco','',15],
    ['Octubre','A','nefro','',32],
    ['Octubre','A','neuro','',20],
    ['Octubre','A','oftalmo','',17],
    ['Octubre','A','onco','',60],
    ['Octubre','A','otras','',96],
    ['Octubre','A','trauma','',43],
    ['Octubre','B','cardio','',1],
    ['Octubre','B','gastro','',3],
    ['Octubre','B','gineco','',3],
    ['Octubre','B','nefro','',1],
    ['Octubre','B','neuro','',1],
    ['Octubre','B','onco','',23],
    ['Octubre','B','otras','',10],
    ['Octubre','B','trauma','',1],
    ['Octubre','C','cardio','',13],
    ['Octubre','C','gastro','',12],
    ['Octubre','C','gineco','',10],
    ['Octubre','C','nefro','',41],
    ['Octubre','C','neuro','',16],
    ['Octubre','C','oftalmo','',65],
    ['Octubre','C','onco','',72],
    ['Octubre','C','otras','',106],
    ['Octubre','C','trauma','',30],
    ['Octubre','C2','cardio','',2],
    ['Octubre','C2','gastro','',3],
    ['Octubre','C2','gineco','',1],
    ['Octubre','C2','nefro','',6],
    ['Octubre','C2','neuro','',2],
    ['Octubre','C2','oftalmo','',2],
    ['Octubre','C2','onco','',8],
    ['Octubre','C2','otras','',11],
    ['Octubre','C2','trauma','',18],
    ['Octubre','D','oftalmo','',4],
    ['Octubre','D','trauma','',1],
    ['Octubre','F','cardio','',3],
    ['Octubre','F','gastro','',3],
    ['Octubre','F','gastro','S',1],
    ['Octubre','F','gineco','',1],
    ['Octubre','F','nefro','',7],
    ['Octubre','F','neuro','',6],
    ['Octubre','F','onco','',11],
    ['Octubre','F','otras','',8],
    ['Octubre','F','otras','S',3],
    ['Octubre','F','trauma','',2],
    ['Octubre','E','cardio','',8],
    ['Octubre','E','cardio','B',1],
    ['Octubre','E','gastro','',12],
    ['Octubre','E','gastro','B',4],
    ['Octubre','E','gineco','',5],
    ['Octubre','E','gineco','B',2],
    ['Octubre','E','nefro','',5],
    ['Octubre','E','nefro','B',3],
    ['Octubre','E','neuro','',3],
    ['Octubre','E','oftalmo','',2],
    ['Octubre','E','onco','',29],
    ['Octubre','E','onco','B',1],
    ['Octubre','E','otras','',22],
    ['Octubre','E','otras','B',8],
    ['Octubre','E','trauma','',19],
    ['Octubre','E','trauma','B',7],
];
const costosInst = [
    // Agosto
    ['Agosto','T','Hospital Privado',2],
    ['Agosto','A','Clinica de Ojos Reyes-Giobellina',15],
    ['Agosto','A','Hospital Italiano',2],
    ['Agosto','A','Hospital Privado',115],
    ['Agosto','A','Instituto Oulton',17],
    ['Agosto','A','Nascentis',1],
    ['Agosto','A','Sanatorio Allende',28],
    ['Agosto','A','Sanatorio Guemes',1],
    ['Agosto','A','Sanatorio Mayo',2],
    ['Agosto','B','Hospital Privado',5],
    ['Agosto','B','Instituto Oulton',32],
    ['Agosto','B','Sanatorio Allende',1],
    ['Agosto','C','Clinica de Ojos Reyes-Giobellina',63],
    ['Agosto','C','Fecundart',2],
    ['Agosto','C','Fund. Bulimia y Anorexia',1],
    ['Agosto','C','Hospital Italiano',1],
    ['Agosto','C','Hospital Privado',227],
    ['Agosto','C','Instituto Oulton',24],
    ['Agosto','C','Instituto Radioterapia',10],
    ['Agosto','C','Sanatorio Allende',56],
    ['Agosto','C','Sanatorio Mayo',3],
    ['Agosto','C','Sanatorio Morra',1],
    ['Agosto','C2','Clinica de Ojos Reyes-Giobellina',3],
    ['Agosto','C2','Hospital Privado',15],
    ['Agosto','C2','Instituto Oulton',1],
    ['Agosto','C2','Sanatorio Allende',6],
    ['Agosto','C2','Sanatorio Mayo',2],
    ['Agosto','D','Clinica de Ojos Reyes-Giobellina',15],
    ['Agosto','D','Hospital Privado',3],
    ['Agosto','D','Instituto Oulton',1],
    ['Agosto','D','Sanatorio Allende',1],
    ['Agosto','F','Clinica de Ojos Reyes-Giobellina',1],
    ['Agosto','F','Hospital Italiano',1],
    ['Agosto','F','Hospital Privado',53],
    ['Agosto','F','Instituto Oulton',4],
    ['Agosto','F','Instituto Radioterapia',1],
    ['Agosto','F','Sanatorio Allende',5],
    ['Agosto','F','Sanatorio Mayo',1],
    ['Agosto','F','Sanatorio Morra',3],
    ['Agosto','E','Clinica de Ojos Reyes-Giobellina',5],
    ['Agosto','E','Hospital Privado',73],
    ['Agosto','E','Instituto Oulton',6],
    ['Agosto','E','Instituto Radioterapia',2],
    ['Agosto','E','Sanatorio Allende',22],
    ['Agosto','E','Sanatorio Mayo',4],
    // Septiembre
    ['Septiembre','T','Hospital Italiano',1],
    ['Septiembre','T','Hospital Privado',1],
    ['Septiembre','A','Centro Kuspide',1],
    ['Septiembre','A','Clinica Central de Ojos',1],
    ['Septiembre','A','Clinica de Ojos Reyes-Giobellina',29],
    ['Septiembre','A','Hospital Alemán',1],
    ['Septiembre','A','Hospital Italiano',1],
    ['Septiembre','A','Hospital Privado',161],
    ['Septiembre','A','Instituto Oulton',14],
    ['Septiembre','A','Instituto Radioterapia',3],
    ['Septiembre','A','Nascentis',2],
    ['Septiembre','A','Sanatorio Allende',35],
    ['Septiembre','A','Sanatorio Guemes',2],
    ['Septiembre','A','Sanatorio Mayo',7],
    ['Septiembre','B','Clinica de Ojos Reyes-Giobellina',1],
    ['Septiembre','B','Fecundart',1],
    ['Septiembre','B','Hospital Privado',10],
    ['Septiembre','B','Instituto Oulton',31],
    ['Septiembre','C','Centro Kuspide',1],
    ['Septiembre','C','Clinica Central de Ojos',1],
    ['Septiembre','C','Clinica de Ojos Reyes-Giobellina',64],
    ['Septiembre','C','Fecundart',4],
    ['Septiembre','C','Hospital Garrahan',1],
    ['Septiembre','C','Hospital Italiano',7],
    ['Septiembre','C','Hospital Privado',214],
    ['Septiembre','C','Instituto Oulton',22],
    ['Septiembre','C','Instituto Radioterapia',6],
    ['Septiembre','C','Nascentis',4],
    ['Septiembre','C','Sanatorio Allende',47],
    ['Septiembre','C','Sanatorio Mayo',15],
    ['Septiembre','C2','Clinica de Ojos Reyes-Giobellina',4],
    ['Septiembre','C2','Hospital Garrahan',1],
    ['Septiembre','C2','Hospital Privado',23],
    ['Septiembre','C2','Instituto Oulton',1],
    ['Septiembre','C2','Sanatorio Allende',18],
    ['Septiembre','C2','Sanatorio Mayo',2],
    ['Septiembre','D','Clinica de Ojos Reyes-Giobellina',7],
    ['Septiembre','F','Clinica de Ojos Reyes-Giobellina',2],
    ['Septiembre','F','Hospital Privado',49],
    ['Septiembre','F','Instituto Oulton',6],
    ['Septiembre','F','Instituto Radioterapia',1],
    ['Septiembre','F','Sanatorio Allende',3],
    ['Septiembre','F','Sanatorio Mayo',4],
    ['Septiembre','F','Sanatorio Morra',1],
    ['Septiembre','E','Clinica de Ojos Reyes-Giobellina',8],
    ['Septiembre','E','Hospital Italiano',1],
    ['Septiembre','E','Hospital Privado',79],
    ['Septiembre','E','Instituto Oulton',3],
    ['Septiembre','E','Instituto Radioterapia',1],
    ['Septiembre','E','Sanatorio Allende',29],
    ['Septiembre','E','Sanatorio Guemes',1],
    ['Septiembre','E','Sanatorio Mayo',1],
    // Octubre
    ['Octubre','T','Hospital Italiano',1],
    ['Octubre','T','Hospital Privado',1],
    ['Octubre','A','Clinica de Ojos Reyes-Giobellina',21],
    ['Octubre','A','Clínica Velez Sarsfield',6],
    ['Octubre','A','Hospital Alemán',3],
    ['Octubre','A','Hospital Italiano',3],
    ['Octubre','A','Hospital Privado',211],
    ['Octubre','A','Instituto Oulton',38],
    ['Octubre','A','Sanatorio Allende',34],
    ['Octubre','A','Sanatorio Guemes',1],
    ['Octubre','A','Sanatorio Mayo',9],
    ['Octubre','A','Sanatorio Morra',1],
    ['Octubre','B','Hospital Privado',8],
    ['Octubre','B','Instituto Oulton',34],
    ['Octubre','B','Sanatorio Allende',1],
    ['Octubre','C','Centro Kuspide',1],
    ['Octubre','C','Clinica Central de Ojos',1],
    ['Octubre','C','Clinica de Ojos Reyes-Giobellina',86],
    ['Octubre','C','Clínica Velez Sarsfield',9],
    ['Octubre','C','Fecundart',2],
    ['Octubre','C','Hospital Garrahan',1],
    ['Octubre','C','Hospital Italiano',3],
    ['Octubre','C','Hospital Privado',185],
    ['Octubre','C','Instituto Oulton',26],
    ['Octubre','C','Instituto Radioterapia',9],
    ['Octubre','C','Nascentis',4],
    ['Octubre','C','Sanatorio Allende',30],
    ['Octubre','C','Sanatorio Mayo',8],
    ['Octubre','C2','Clinica de Ojos Reyes-Giobellina',2],
    ['Octubre','C2','Hospital Italiano',1],
    ['Octubre','C2','Hospital Privado',29],
    ['Octubre','C2','Instituto Oulton',4],
    ['Octubre','C2','Sanatorio Allende',16],
    ['Octubre','C2','Sanatorio Mayo',1],
    ['Octubre','D','Clinica de Ojos Reyes-Giobellina',4],
    ['Octubre','D','Hospital Privado',1],
    ['Octubre','F','Clinica de Ojos Reyes-Giobellina',1],
    ['Octubre','F','Clínica Velez Sarsfield',5],
    ['Octubre','F','Hospital Alemán',1],
    ['Octubre','F','Hospital Privado',25],
    ['Octubre','F','Instituto Oulton',1],
    ['Octubre','F','Instituto Radioterapia',2],
    ['Octubre','F','Sanatorio Allende',10],
    ['Octubre','E','Clinica de Ojos Reyes-Giobellina',3],
    ['Octubre','E','Clínica Velez Sarsfield',2],
    ['Octubre','E','Hospital Garrahan',1],
    ['Octubre','E','Hospital Privado',95],
    ['Octubre','E','Instituto Oulton',6],
    ['Octubre','E','Nascentis',2],
    ['Octubre','E','Sanatorio Allende',19],
    ['Octubre','E','Sanatorio Guemes',1],
    ['Octubre','E','Sanatorio Mayo',2],
];
const costosOrig = [
    // Agosto
    ['Agosto','T','LR',2],
    ['Agosto','A','CBA',12],
    ['Agosto','A','LR',169],
    ['Agosto','B','LR',38],
    ['Agosto','C','CBA',59],
    ['Agosto','C','LR',329],
    ['Agosto','C2','LR',27],
    ['Agosto','D','LR',20],
    ['Agosto','F','CBA',32],
    ['Agosto','F','LR',37],
    ['Agosto','E','CBA',41],
    ['Agosto','E','LR',71],
    // Septiembre
    ['Septiembre','T','LR',2],
    ['Septiembre','A','CBA',15],
    ['Septiembre','A','LR',241],
    ['Septiembre','A','SIN',1],
    ['Septiembre','B','CBA',1],
    ['Septiembre','B','LR',42],
    ['Septiembre','C','CBA',69],
    ['Septiembre','C','LR',310],
    ['Septiembre','C','SIN',7],
    ['Septiembre','C2','LR',49],
    ['Septiembre','D','CBA',1],
    ['Septiembre','D','LR',6],
    ['Septiembre','F','CBA',27],
    ['Septiembre','F','LR',38],
    ['Septiembre','F','SIN',1],
    ['Septiembre','E','CBA',37],
    ['Septiembre','E','LR',86],
    // Octubre
    ['Octubre','T','LR',2],
    ['Octubre','A','CBA',14],
    ['Octubre','A','LR',313],
    ['Octubre','B','CBA',2],
    ['Octubre','B','LR',41],
    ['Octubre','C','CBA',59],
    ['Octubre','C','LR',306],
    ['Octubre','C2','LR',53],
    ['Octubre','D','LR',5],
    ['Octubre','F','CBA',13],
    ['Octubre','F','LR',32],
    ['Octubre','E','CBA',28],
    ['Octubre','E','LR',103],
];

const costosMeses = () => MESES_ORDEN.filter(m => costosCat.some(r => r[0]===m));
const cSel   = arr => costosPeriodo==='Total' ? arr : arr.filter(r => r[0]===costosPeriodo);
const cNum   = v => v.toLocaleString('es-AR');
const cPct   = v => (isFinite(v)?v:0).toFixed(1).replace('.',',')+'%';
function cFmt(v) {
    if (costosUnidad==='pesos') return '$ '+Math.round(v).toLocaleString('es-AR');
    return v.toLocaleString('es-AR',{minimumFractionDigits:1,maximumFractionDigits:1});
}
function cGroupBy(rows, keyFn) {
    const m = {};
    rows.forEach(r => {
        const n = r[r.length-1], g = r[1], k = keyFn(r);
        const o = m[k] || (m[k] = {n:0, c:0, int:0, g:{}});
        o.n += n; o.c += n*(costosPesos[g]||0); if (g==='E'||g==='F') o.int += n;
        o.g[g] = (o.g[g]||0) + n;
    });
    return m;
}

function setCostosPeriodo(p) { costosPeriodo = p; renderCostos(); }
function setCostosUnidad(u) { costosUnidad = u; renderCostos(); }
function setCostosPeso(k, v) { const x = parseFloat(String(v).replace(',','.')); costosPesos[k] = isFinite(x)&&x>=0 ? x : 0; renderCostos(false); }
function resetCostosPesos() { Object.assign(costosPesos, COSTOS_PESOS_DEF); renderCostos(); }

function setVista(v) {
    ['monitor','costos'].forEach(x => {
        document.getElementById('view-'+x).classList.toggle('hidden', x!==v);
        const b = document.getElementById('btn-vista-'+x);
        b.classList.toggle('tab-btn-active', x===v); b.classList.toggle('tab-btn-inactive', x!==v);
    });
    document.getElementById('nav-filtros').classList.toggle('hidden', v!=='monitor');
    if (v==='costos') renderCostos(); else procesarDatos();
    window.scrollTo({top:0});
}

function renderCostos(rebuildPesos = true) {
    const meses = costosMeses();
    if (costosPeriodo!=='Total' && !meses.includes(costosPeriodo)) costosPeriodo = 'Total';
    document.getElementById('costos-meses-cubiertos').innerText = meses.join(', ');
    document.getElementById('costos-subtitulo').innerText =
        `${costosPeriodo==='Total' ? meses[0]+' a '+meses[meses.length-1]+' 2026' : costosPeriodo+' 2026'} · costo en ${costosUnidad==='pesos'?'pesos ($)':'índice relativo (valores supuestos)'}`;

    // selector de período
    document.getElementById('costos-periodos').innerHTML = ['Total',...meses].map(p =>
        `<button onclick="setCostosPeriodo('${p}')" class="tab-btn ${p===costosPeriodo?'tab-btn-active':'tab-btn-inactive'}">${p==='Total'?'Total período':p}</button>`).join('');
    // editor de pesos
    if (rebuildPesos) {
        document.getElementById('costos-unidad').value = costosUnidad;
        document.getElementById('costos-pesos').innerHTML = COSTOS_GRUPOS.map(g => `
            <label class="block">
                <span class="flex items-center gap-2 text-[11px] text-slate-500 mb-1"><span style="width:8px;height:8px;border-radius:50%;background:${g.color};display:inline-block"></span>${g.label}</span>
                <input type="number" min="0" step="any" value="${costosPesos[g.k]}" onchange="setCostosPeso('${g.k}',this.value)"
                    class="w-full bg-white border border-slate-200 rounded-md px-3 py-1.5 text-sm font-normal text-slate-500 focus:outline-none focus:ring-2 focus:ring-aposTeal">
            </label>`).join('');
    }

    const cat = cSel(costosCat);
    const byG = cGroupBy(cat, r => r[1]);
    const totN = Object.values(byG).reduce((s,o)=>s+o.n,0);
    const totC = Object.values(byG).reduce((s,o)=>s+o.c,0);
    const gN = k => (byG[k]?byG[k].n:0), gC = k => (byG[k]?byG[k].c:0);
    const intN = gN('E')+gN('F'), intC = gC('E')+gC('F');

    // KPIs
    document.getElementById('cs-total').innerText = cNum(totN);
    document.getElementById('cs-total-sub').innerText = costosPeriodo==='Total' ? `${meses.length} meses con detalle cargado` : 'derivaciones del mes';
    document.getElementById('cs-int').innerText = cPct(totN?intN/totN*100:0);
    document.getElementById('cs-int-sub').innerText = `${cNum(intN)} casos · ${cNum(gN('E'))} con cirugía, ${cNum(gN('F'))} clínicas`;
    document.getElementById('cs-intcosto').innerText = cPct(totC?intC/totC*100:0);
    document.getElementById('cs-intcosto-sub').innerText = `con el ${cPct(totN?intN/totN*100:0)} de los casos`;
    document.getElementById('cs-medio').innerText = totN ? cFmt(totC/totN) : '—';
    document.getElementById('cs-medio-sub').innerText = `costo total estimado: ${cFmt(totC)}`;

    // Gráfico casos vs costo por tipo
    const gs = COSTOS_GRUPOS.filter(g => gN(g.k)>0);
    if (chartCostoTipoObj) chartCostoTipoObj.destroy();
    chartCostoTipoObj = new Chart(document.getElementById('chartCostoTipo').getContext('2d'), {
        type:'bar',
        data:{ labels: gs.map(g=>g.label), datasets:[
            { label:'% de casos', data: gs.map(g=>totN?gN(g.k)/totN*100:0), backgroundColor:'#45b5a9', borderRadius:4, categoryPercentage:0.8, barPercentage:0.9 },
            { label:'% del costo estimado', data: gs.map(g=>totC?gC(g.k)/totC*100:0), backgroundColor:'#f08a8a', borderRadius:4, categoryPercentage:0.8, barPercentage:0.9 }
        ]},
        options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, layout:{padding:{right:38}},
            scales:{ x:{beginAtZero:true, ticks:{callback:v=>v+'%', font:{size:10}}, grid:{color:'#f1f5f9'}}, y:{ticks:{font:{size:10}}, grid:{display:false}} },
            plugins:{ legend:{position:'top',labels:{usePointStyle:true,boxWidth:8,padding:16}},
                datalabels:{anchor:'end',align:'right',color:'#6b7280',font:{weight:'400',size:9},formatter:v=>v.toFixed(1).replace('.',',')+'%',clip:false} } }
    });

    // Evolución mensual (siempre todos los meses)
    const evo = meses.map(m => {
        const o = cGroupBy(costosCat.filter(r=>r[0]===m), r=>'x').x || {n:0,c:0,int:0};
        return {m, n:o.n, intPct:o.n?o.int/o.n*100:0, medio:o.n?o.c/o.n:0};
    });
    if (chartCostoMesObj) chartCostoMesObj.destroy();
    chartCostoMesObj = new Chart(document.getElementById('chartCostoMes').getContext('2d'), {
        type:'bar',
        data:{ labels: evo.map(e=>e.m), datasets:[
            { type:'bar',  label:'Derivaciones', data:evo.map(e=>e.n), backgroundColor:'#b9d9d5', borderRadius:4, yAxisID:'y',
              datalabels:{anchor:'end',align:'top',color:'#6b7280',font:{size:9},formatter:v=>cNum(v)} },
            { type:'line', label:'Costo medio / derivación', data:evo.map(e=>e.medio), borderColor:'#f08a8a', backgroundColor:'#f08a8a', tension:0.3, pointRadius:4, yAxisID:'y2',
              datalabels:{anchor:'end',align:'top',color:'#f08a8a',font:{size:9},formatter:v=>cFmt(v)} }
        ]},
        options:{ responsive:true, maintainAspectRatio:false, layout:{padding:{top:22}},
            scales:{ y:{beginAtZero:true, grace:'15%', ticks:{font:{size:10}}, grid:{color:'#f1f5f9'}},
                     y2:{position:'right', beginAtZero:false, grace:'20%', display:false},
                     x:{ticks:{font:{size:10}}, grid:{display:false}} },
            plugins:{ legend:{position:'top',labels:{usePointStyle:true,boxWidth:8,padding:12}} } }
    });

    // Tabla por diagnóstico
    const hdr = (cols) => `<thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}</tr></thead>`;
    const bar = (p,color) => `<div class="cost-bar"><span style="width:${Math.min(100,p)}%;background:${color}"></span></div>`;
    const byCat = cGroupBy(cat, r=>r[2]);
    const catRows = Object.entries(byCat).sort((a,b)=>b[1].c-a[1].c).map(([k,o]) => {
        const surg = (o.g.E||0), amb = (o.g.C||0);
        return `<tr><td>${COSTOS_CAT[k]||k}</td><td class="num">${cNum(o.n)}</td><td class="num">${cPct(o.n/totN*100)}</td>
            <td class="num">${cPct(o.int/o.n*100)}</td><td class="num">${cPct(surg/o.n*100)}</td><td class="num">${cPct(amb/o.n*100)}</td>
            <td class="num">${cFmt(o.c)}</td><td>${bar(totC?o.c/totC*100:0,'#f08a8a')}<span class="cost-pct">${cPct(totC?o.c/totC*100:0)}</span></td><td class="num">${cFmt(o.c/o.n)}</td></tr>`;
    }).join('');
    document.getElementById('tabla-costo-dx').innerHTML = hdr(['Diagnóstico','Casos','% casos','% internación','% int. con cirugía','% trat. ambulatorio','Costo estimado','% del costo','Costo medio']) + `<tbody>${catRows}</tbody>`;

    // Tabla por prestador
    const inst = cGroupBy(cSel(costosInst), r=>r[2]);
    const instRows = Object.entries(inst).sort((a,b)=>b[1].c-a[1].c).slice(0,10).map(([k,o]) =>
        `<tr><td>${k}</td><td class="num">${cNum(o.n)}</td><td class="num">${cPct(o.n/totN*100)}</td><td class="num">${cPct(o.int/o.n*100)}</td>
         <td>${bar(totC?o.c/totC*100:0,'#45b5a9')}<span class="cost-pct">${cPct(totC?o.c/totC*100:0)}</span></td><td class="num">${cFmt(o.c/o.n)}</td></tr>`).join('');
    document.getElementById('tabla-costo-inst').innerHTML = hdr(['Prestador','Casos','% casos','% internación','% del costo','Costo medio']) + `<tbody>${instRows}</tbody>`;

    // Tabla por origen
    const org = cGroupBy(cSel(costosOrig), r=>r[2]);
    const orgRows = ['LR','CBA','SIN'].filter(k=>org[k]).map(k => { const o = org[k];
        return `<tr><td>${COSTOS_ORG[k]}</td><td class="num">${cNum(o.n)}</td><td class="num">${cPct(o.int/o.n*100)}</td><td class="num">${cPct(totC?o.c/totC*100:0)}</td><td class="num">${cFmt(o.c/o.n)}</td></tr>`; }).join('');
    document.getElementById('tabla-costo-org').innerHTML = hdr(['Origen','Casos','% internación','% del costo','Costo medio']) + `<tbody>${orgRows}</tbody>`;

    // Lectura y oportunidades (se recalcula con los datos del período)
    const ins = [];
    ins.push({i:'fa-bed', c:'#f08a8a', t:`La internación (clínica + cirugía) es el <span class="font-medium text-slate-500">${cPct(totN?intN/totN*100:0)}</span> de los casos y concentra el <span class="font-medium text-slate-500">${cPct(totC?intC/totC*100:0)}</span> del costo estimado. Con la cirugía ambulatoria llega a ${cPct(totC?(intC+gC('D'))/totC*100:0)}.`});
    const benN = cat.filter(r=>r[3]==='B').reduce((s,r)=>s+r[4],0);
    if (gN('E')) ins.push({i:'fa-calendar-check', c:'#45b5a9', t:`<span class="font-medium text-slate-500">${cNum(benN)}</span> de ${cNum(gN('E'))} internaciones con cirugía (${cPct(benN/gN('E')*100)}) son de patología benigna o electiva (miomatosis, endometriosis, litiasis, hemorroides, hernias, cataratas, artrosis…). Son programables y negociables por módulo; evaluar resolución local o convenios.`});
    const sinN = cat.filter(r=>r[3]==='S').reduce((s,r)=>s+r[4],0);
    if (gN('F')) ins.push({i:'fa-magnifying-glass', c:'#f2bf78', t:`<span class="font-medium text-slate-500">${cNum(sinN)}</span> de ${cNum(gN('F'))} internaciones clínicas (${cPct(sinN/gN('F')*100)}) tienen diagnóstico sintomático o "en estudio" (dolor precordial, síncope, anemia en estudio…). Candidatas a hospital de día o estudio ambulatorio, previa validación médica.`});
    const instArr = Object.entries(inst).sort((a,b)=>b[1].n-a[1].n);
    if (instArr.length) { const [k,o] = instArr[0];
        const mayor = Object.entries(inst).filter(([,x])=>x.n>=20).sort((a,b)=>b[1].c/b[1].n-a[1].c/a[1].n)[0];
        ins.push({i:'fa-hospital', c:'#92bde4', t:`<span class="font-medium text-slate-500">${k}</span> reúne el ${cPct(o.n/totN*100)} de los casos y el ${cPct(totC?o.c/totC*100:0)} del costo estimado.${mayor?` El mayor costo medio entre prestadores con 20+ casos es <span class="font-medium text-slate-500">${mayor[0]}</span> (${cFmt(mayor[1].c/mayor[1].n)}).`:''} Priorizar auditoría de módulos y prótesis en estos prestadores.`}); }
    if (org.CBA && org.LR) ins.push({i:'fa-route', c:'#ab9fbb', t:`Las derivaciones con origen en Córdoba son el ${cPct(org.CBA.n/totN*100)} de los casos y el ${cPct(totC?org.CBA.c/totC*100:0)} del costo; se internan el ${cPct(org.CBA.int/org.CBA.n*100)} contra ${cPct(org.LR.int/org.LR.n*100)} de las de La Rioja.`});
    if (costosPeriodo==='Total' && evo.length>1) { const a=evo[0], z=evo[evo.length-1];
        ins.push({i:'fa-chart-line', c:'#7dd1a0', t:`${a.m} → ${z.m}: la internación pasó de ${cPct(a.intPct)} a ${cPct(z.intPct)} de los casos y el costo medio por derivación de ${cFmt(a.medio)} a ${cFmt(z.medio)} (${((z.medio/a.medio-1)*100).toFixed(1).replace('.',',')}%), con ${cNum(a.n)} → ${cNum(z.n)} derivaciones.`}); }
    document.getElementById('costos-insights').innerHTML = ins.map(x => `
        <div class="insight-card"><i class="fas ${x.i}" style="color:${x.c}"></i><p>${x.t}</p></div>`).join('');
}

window.onload = () => procesarDatos();
