/*
MySQL Backup
Source Server Version: 5.1.30
Source Database: ferreteria
Date: 24/08/2019 18:30:08
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
--  View definition for `vegreso`
-- ----------------------------
DROP VIEW IF EXISTS `vegreso`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vegreso` AS select `cabegreso`.`cegr_fec` AS `cegr_fec`,`cabegreso`.`cegr_obs` AS `cegr_obs`,`cabegreso`.`cegr_id` AS `cegr_id`,`cabegreso`.`Usu_id` AS `Usu_id`,`detegreso`.`degr_id` AS `degr_id`,`detegreso`.`degr_can` AS `degr_can`,`ferreteria`.`detegreso`.`degr_pu` AS `degr_pu`,`ferreteria`.`detegreso`.`degr_tot` AS `degr_tot`,`ferreteria`.`detegreso`.`degr_ppp` AS `degr_ppp`,`ferreteria`.`productos`.`pro_id` AS `pro_id`,`ferreteria`.`productos`.`pro_Des` AS `pro_Des`,`ferreteria`.`productos`.`Pro_uni` AS `Pro_uni`,`ferreteria`.`productos`.`pro_tip` AS `pro_tip`,`ferreteria`.`productos`.`prv_id` AS `prv_id`,`ferreteria`.`productos`.`pro_exmi` AS `pro_exmi`,`ferreteria`.`productos`.`pro_exma` AS `pro_exma`,`ferreteria`.`productos`.`pro_ubi` AS `pro_ubi`,`ferreteria`.`productos`.`pro_cod` AS `pro_cod`,`ferreteria`.`productos`.`pro_est` AS `pro_est`,`ferreteria`.`productos`.`pro_aut` AS `pro_aut`,`ferreteria`.`productos`.`mar_id` AS `mar_id`,`ferreteria`.`productos`.`gru_id` AS `gru_id`,`ferreteria`.`marcas`.`Mar_des` AS `Mar_des` from (((`cabegreso` join `detegreso` on((`ferreteria`.`cabegreso`.`cegr_id` = `ferreteria`.`detegreso`.`cegr_id`))) join `productos` on((`ferreteria`.`detegreso`.`pro_id` = `ferreteria`.`productos`.`pro_id`))) join `marcas` on((`ferreteria`.`productos`.`mar_id` = `ferreteria`.`marcas`.`Mar_id`)));

-- ----------------------------
--  View definition for `vingreso`
-- ----------------------------
DROP VIEW IF EXISTS `vingreso`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vingreso` AS select `cabingreso`.`cingr_id` AS `cingr_id`,`detingreso`.`ding_can` AS `ding_can`,`detingreso`.`ding_prt` AS `ding_prt`,`detingreso`.`pro_prv` AS `pro_prv`,`detingreso`.`ding_pru` AS `ding_pru`,`productos`.`pro_Des` AS `pro_Des`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`pro_cod` AS `pro_cod`,`productos`.`pro_exi` AS `pro_exi` from ((`cabingreso` join `detingreso` on((`cabingreso`.`cingr_id` = `detingreso`.`cing_Id`))) join `productos` on((`detingreso`.`pro_id` = `productos`.`pro_id`)));

-- ----------------------------
--  View definition for `vproducto`
-- ----------------------------
DROP VIEW IF EXISTS `vproducto`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vproducto` AS select `marcas`.`Mar_des` AS `Mar_des`,`vproprv`.`pro_Des` AS `pro_Des`,`vproprv`.`Prv_des` AS `Prv_des`,`vproprv`.`Pro_uni` AS `Pro_uni`,`vproprv`.`gru_id` AS `gru_id`,`vproprv`.`pro_cod` AS `pro_cod`,`vproprv`.`mar_id` AS `mar_id`,`grupos`.`gru_des` AS `gru_des` from ((`vproprv` join `grupos` on((`vproprv`.`gru_id` = `grupos`.`gru_id`))) join `marcas` on((`vproprv`.`mar_id` = `marcas`.`Mar_id`)));

-- ----------------------------
--  View definition for `vproductos`
-- ----------------------------
DROP VIEW IF EXISTS `vproductos`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vproductos` AS select `proveedores`.`Prv_des` AS `Prv_des`,`productos`.`pro_Des` AS `pro_Des`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`pro_cod` AS `pro_cod`,`marcas`.`Mar_des` AS `Mar_des`,`grupos`.`gru_des` AS `gru_des` from (((`productos` join `proveedores` on((`productos`.`prv_id` = `proveedores`.`Prv_id`))) join `marcas` on((`productos`.`mar_id` = `marcas`.`Mar_id`))) join `grupos` on((`productos`.`gru_id` = `grupos`.`gru_id`)));

-- ----------------------------
--  View definition for `vproprv`
-- ----------------------------
DROP VIEW IF EXISTS `vproprv`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vproprv` AS select `productos`.`pro_Des` AS `pro_Des`,`proveedores`.`Prv_des` AS `Prv_des`,`productos`.`Pro_uni` AS `Pro_uni`,`productos`.`gru_id` AS `gru_id`,`productos`.`pro_cod` AS `pro_cod`,`productos`.`mar_id` AS `mar_id`,`proprv`.`prv_id` AS `prv_id`,`proprv`.`pro_id` AS `pro_id` from ((`proprv` join `productos` on((`proprv`.`pro_id` = `productos`.`pro_id`))) join `proveedores` on((`proprv`.`prv_id` = `proveedores`.`Prv_id`)));

-- ----------------------------
--  Records 
-- ----------------------------
