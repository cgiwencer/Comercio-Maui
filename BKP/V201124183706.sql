/*
MySQL Backup
Source Server Version: 5.1.30
Source Database: comercio
Date: 24/11/2020 18:37:06
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
--  View definition for `valming`
-- ----------------------------
DROP VIEW IF EXISTS `valming`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `valming` AS select `cabinvinicial`.`ini_id` AS `ini_id`,`cabinvinicial`.`ini_fec` AS `ini_fec`,`cabinvinicial`.`almid` AS `almid`,`almacen`.`AlmDes` AS `AlmDes` from (`cabinvinicial` join `almacen` on((`cabinvinicial`.`almid` = `almacen`.`AlmId`)));

-- ----------------------------
--  View definition for `varqueo`
-- ----------------------------
DROP VIEW IF EXISTS `varqueo`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `varqueo` AS select `usuarios`.`Usu_Nom` AS `Usu_Nom`,`arqueo`.`Pag_fec` AS `Pag_fec`,`arqueo`.`Pag_NFa` AS `Pag_NFa`,`arqueo`.`Pag_Nit` AS `Pag_Nit`,`arqueo`.`Pag_RaS` AS `Pag_RaS`,`arqueo`.`Pag_Sut` AS `Pag_Sut`,`arqueo`.`Pag_DBs` AS `Pag_DBs`,`arqueo`.`Pag_Dpo` AS `Pag_Dpo`,`arqueo`.`Pag_Mon` AS `Pag_Mon` from (`arqueo` join `usuarios` on((`arqueo`.`Usu_Id` = `usuarios`.`Usu_Id`)));

-- ----------------------------
--  View definition for `vcompra`
-- ----------------------------
DROP VIEW IF EXISTS `vcompra`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vcompra` AS select `cabingreso`.`cingr_fec` AS `cingr_fec`,`detingreso`.`pro_cod` AS `pro_cod`,`detingreso`.`pro_des` AS `pro_des`,`detingreso`.`ding_can` AS `ding_can`,`cabingreso`.`cing_obs` AS `cing_obs` from (`cabingreso` join `detingreso` on((`cabingreso`.`cingr_id` = `detingreso`.`cing_Id`)));

-- ----------------------------
--  View definition for `vegresos`
-- ----------------------------
DROP VIEW IF EXISTS `vegresos`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vegresos` AS select `cabegreso`.`cegr_fec` AS `cegr_fec`,`detegreso`.`pro_cod` AS `pro_cod`,`detegreso`.`pro_des` AS `pro_des`,`detegreso`.`degr_can` AS `degr_can`,`detegreso`.`degr_pru` AS `degr_pru`,`detegreso`.`degr_prt` AS `degr_prt`,`cabegreso`.`cegr_id` AS `cegr_id` from (`cabegreso` join `detegreso` on((`cabegreso`.`cegr_id` = `detegreso`.`cegr_id`)));

-- ----------------------------
--  View definition for `vingreso`
-- ----------------------------
DROP VIEW IF EXISTS `vingreso`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vingreso` AS select `cabingreso`.`cingr_id` AS `cingr_id`,`detingreso`.`ding_can` AS `ding_can`,`detingreso`.`ding_prt` AS `ding_prt`,`detingreso`.`pro_prv` AS `pro_prv`,`detingreso`.`ding_pru` AS `ding_pru`,`productos`.`pro_Des` AS `pro_Des`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`pro_cod` AS `pro_cod`,`productos`.`pro_exi` AS `pro_exi` from ((`cabingreso` join `detingreso` on((`cabingreso`.`cingr_id` = `detingreso`.`cing_Id`))) join `productos` on((`detingreso`.`pro_id` = `productos`.`pro_id`)));

-- ----------------------------
--  View definition for `vinvini`
-- ----------------------------
DROP VIEW IF EXISTS `vinvini`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vinvini` AS select `cabinvinicial`.`ini_fec` AS `ini_fec`,`detinvinicial`.`Pro_cod` AS `Pro_cod`,`detinvinicial`.`Pro_des` AS `Pro_des`,`detinvinicial`.`Ini_Can` AS `Ini_Can` from (`cabinvinicial` join `detinvinicial` on((`cabinvinicial`.`ini_id` = `detinvinicial`.`InI_Id`)));

-- ----------------------------
--  View definition for `vpro1`
-- ----------------------------
DROP VIEW IF EXISTS `vpro1`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vpro1` AS select `productos`.`pro_id` AS `pro_id`,`productos`.`pro_Des` AS `pro_Des`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`pro_tip` AS `pro_tip`,`productos`.`pro_cod` AS `pro_cod`,`productos`.`pro_ubi` AS `pro_ubi`,`grupos`.`gru_des` AS `gru_des`,`productos`.`Pro_sal` AS `Pro_sal`,`productos`.`ProUni` AS `ProUni` from (`productos` join `grupos` on((`productos`.`gru_id` = `grupos`.`gru_id`)));

-- ----------------------------
--  View definition for `vprodbaj`
-- ----------------------------
DROP VIEW IF EXISTS `vprodbaj`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vprodbaj` AS select `productos`.`pro_Des` AS `pro_Des`,`bajas`.`Baj_fec` AS `Baj_fec`,`bajas`.`Pro_cod` AS `Pro_cod`,`bajas`.`Baj_can` AS `Baj_can`,`bajas`.`Baj_mot` AS `Baj_mot`,`usuarios`.`Usu_Nom` AS `Usu_Nom`,`bajas`.`Baj_est` AS `Baj_est`,`bajas`.`cegr_id` AS `cegr_id`,`bajas`.`Baj_id` AS `Baj_id` from ((`bajas` join `productos` on((`bajas`.`Pro_cod` = `productos`.`pro_cod`))) join `usuarios` on((`bajas`.`Usu_Id` = `usuarios`.`Usu_Id`)));

-- ----------------------------
--  View definition for `vproducto`
-- ----------------------------
DROP VIEW IF EXISTS `vproducto`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vproducto` AS select `grupos`.`gru_des` AS `gru_des`,`productos`.`pro_Des` AS `pro_Des`,`productos`.`pro_tip` AS `pro_tip`,`productos`.`pro_ubi` AS `pro_ubi`,`productos`.`pro_cod` AS `pro_cod`,`productos`.`Pro_sal` AS `Pro_sal`,`marcas`.`Mar_des` AS `Mar_des`,`productos`.`gru_id` AS `gru_id`,`productos`.`pro_id` AS `pro_id`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`pro_exmi` AS `pro_exmi`,`productos`.`pro_exma` AS `pro_exma`,`productos`.`mar_id` AS `mar_id`,`productos`.`pro_est` AS `pro_est`,`productos`.`pro_aut` AS `pro_aut`,`productos`.`pro_PrC` AS `pro_PrC`,`productos`.`prv_id` AS `prv_id`,`productos`.`pro_pve` AS `pro_pve`,`productos`.`pro_exi` AS `pro_exi`,`productos`.`pro_ppp` AS `pro_ppp`,`productos`.`Usu_id` AS `Usu_id`,`productos`.`ProUni` AS `ProUni`,`productos`.`ProTNu` AS `ProTNu`,`productos`.`ProTLi` AS `ProTLi`,`productos`.`pro_codb` AS `pro_codb` from ((`productos` join `grupos` on((`grupos`.`gru_id` = `productos`.`gru_id`))) join `marcas` on((`productos`.`mar_id` = `marcas`.`Mar_id`)));

-- ----------------------------
--  View definition for `vproductos`
-- ----------------------------
DROP VIEW IF EXISTS `vproductos`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vproductos` AS select `proveedores`.`Prv_des` AS `Prv_des`,`productos`.`pro_Des` AS `pro_Des`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`pro_cod` AS `pro_cod`,`grupos`.`gru_des` AS `gru_des` from ((`productos` join `proveedores` on((`productos`.`prv_id` = `proveedores`.`Prv_id`))) join `grupos` on((`productos`.`gru_id` = `grupos`.`gru_id`)));

-- ----------------------------
--  View definition for `vproprv`
-- ----------------------------
DROP VIEW IF EXISTS `vproprv`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vproprv` AS select `productos`.`pro_Des` AS `pro_Des`,`proveedores`.`Prv_des` AS `Prv_des`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`gru_id` AS `gru_id`,`productos`.`pro_cod` AS `pro_cod`,`productos`.`mar_id` AS `mar_id`,`proprv`.`prv_id` AS `prv_id`,`proprv`.`pro_id` AS `pro_id`,`productos`.`pro_pve` AS `pro_pve`,`productos`.`pro_exi` AS `pro_exi`,`productos`.`pro_ppp` AS `pro_ppp`,`productos`.`pro_PrC` AS `pro_PrC` from ((`proprv` join `productos` on((`proprv`.`pro_id` = `productos`.`pro_id`))) join `proveedores` on((`proprv`.`prv_id` = `proveedores`.`Prv_id`)));

-- ----------------------------
--  View definition for `vsalida`
-- ----------------------------
DROP VIEW IF EXISTS `vsalida`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vsalida` AS select `cabegreso`.`cegr_fec` AS `cegr_fec`,`detegreso`.`pro_cod` AS `pro_cod`,`detegreso`.`pro_des` AS `pro_des`,`detegreso`.`degr_can` AS `degr_can`,`detegreso`.`degr_est` AS `degr_est`,`cabegreso`.`cegr_obs` AS `cegr_obs` from (`cabegreso` join `detegreso` on((`cabegreso`.`cegr_id` = `detegreso`.`cegr_id`)));

-- ----------------------------
--  View definition for `vventa`
-- ----------------------------
DROP VIEW IF EXISTS `vventa`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vventa` AS select `vegresos`.`cegr_fec` AS `cegr_fec`,`vegresos`.`pro_cod` AS `pro_cod`,`vegresos`.`pro_des` AS `pro_des`,`vegresos`.`degr_can` AS `degr_can`,`vegresos`.`degr_pru` AS `degr_pru`,`vegresos`.`degr_prt` AS `degr_prt`,`vegresos`.`cegr_id` AS `cegr_id` from (`arqueo` join `vegresos` on((`arqueo`.`cegr_id` = `vegresos`.`cegr_id`)));

-- ----------------------------
--  View definition for `vventa1`
-- ----------------------------
DROP VIEW IF EXISTS `vventa1`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vventa1` AS select `pagoventa`.`Pag_fec` AS `Pag_fec`,`pagoventa`.`Pag_NFa` AS `Pag_NFa`,`pagoventa`.`Pag_Nit` AS `Pag_Nit`,`pagoventa`.`Pag_RaS` AS `Pag_RaS`,`pagoventa`.`Pag_Sut` AS `Pag_Sut`,`pagoventa`.`Pag_DBs` AS `Pag_DBs`,`pagoventa`.`Pag_Dpo` AS `Pag_Dpo`,`pagoventa`.`Pag_Mon` AS `Pag_Mon`,`detegreso`.`pro_des` AS `pro_des`,`detegreso`.`degr_can` AS `degr_can`,`detegreso`.`pro_cod` AS `pro_cod`,`detegreso`.`degr_pru` AS `degr_pru`,`detegreso`.`degr_prt` AS `degr_prt`,`pagoventa`.`cegr_id` AS `cegr_id` from (`pagoventa` join `detegreso` on((`pagoventa`.`cegr_id` = `detegreso`.`cegr_id`)));

-- ----------------------------
--  Records 
-- ----------------------------
