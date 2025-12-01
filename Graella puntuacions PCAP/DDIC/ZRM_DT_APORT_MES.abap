@EndUserText.label : 'Aportacions i mesures'
define table zrm_dt_aport_mes {
  key docclass          : docid;
  key objectid          : docid;
  key z_num_ofer        : z_num_ofer;
  key z_num_aport       : integer;
  z_descripcio          : string length 255;
  z_puntuacio           : z_punts;
  z_data_creacio        : datum;
  z_usuari_creacio      : uname;
  z_data_modif          : datum;
  z_usuari_modif        : uname;
}