/*
SQLyog Community Edition- MySQL GUI v6.03
Host - 5.0.67-community-nt : Database - lgvha
*********************************************************************
Server version : 5.0.67-community-nt
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

create database if not exists `lgvha`;

USE `lgvha`;

/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;

/*Table structure for table `camera` */

CREATE TABLE `camera` (
  `IPAddress` varchar(50) default NULL,
  `CamPort` double default NULL,
  `CamUser` varchar(50) default NULL,
  `CamPwd` varchar(50) default NULL,
  `CamEntrance` int(11) default NULL,
  `CamExit` int(11) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `entrance` */

CREATE TABLE `entrance` (
  `IDNumber` int(20) NOT NULL auto_increment,
  `ACNumber` int(11) NOT NULL,
  `LogTime` datetime NOT NULL,
  `RefFile` varchar(200) default NULL,
  `LogState` varchar(11) default NULL,
  PRIMARY KEY  (`ACNumber`,`LogTime`),
  UNIQUE KEY `IDNumber` (`IDNumber`)
) ENGINE=MyISAM AUTO_INCREMENT=76 DEFAULT CHARSET=latin1;

/*Table structure for table `homeowners` */

CREATE TABLE `homeowners` (
  `ACNumber` int(11) default NULL,
  `Age` int(11) default NULL,
  `Address` varchar(250) default NULL,
  `Telephone` varchar(250) default NULL,
  `Gender` varchar(250) default NULL,
  `LName` varchar(250) default NULL,
  `Fname` varchar(250) default NULL,
  `MName` varchar(250) default NULL,
  `Notes` varchar(250) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `readers` */

CREATE TABLE `readers` (
  `IPAddressEntrance` varchar(20) default NULL,
  `CommKeyEntrance` varchar(20) default NULL,
  `IPAddressExit` varchar(20) default NULL,
  `CommKeyExit` varchar(20) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
