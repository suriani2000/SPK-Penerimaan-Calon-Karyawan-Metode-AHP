-- phpMyAdmin SQL Dump
-- version 3.5.2.2
-- http://www.phpmyadmin.net
--
-- Inang: 127.0.0.1
-- Waktu pembuatan: 21 Jun 2020 pada 07.32
-- Versi Server: 5.5.27
-- Versi PHP: 5.4.7

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Basis data: `db_calonkaryawan`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `batasan`
--

CREATE TABLE IF NOT EXISTS `batasan` (
  `id` int(5) NOT NULL,
  `diterima` decimal(5,2) NOT NULL,
  `ditolak` decimal(5,2) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `batasan`
--

INSERT INTO `batasan` (`id`, `diterima`, `ditolak`) VALUES
(1, 74.00, 75.00);

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_hasilpenilaian`
--

CREATE TABLE IF NOT EXISTS `tbl_hasilpenilaian` (
  `idpenilaian` int(5) NOT NULL,
  `nilaiakhir` decimal(5,2) NOT NULL,
  `keteranganhasil` varchar(50) NOT NULL,
  PRIMARY KEY (`idpenilaian`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tbl_hasilpenilaian`
--

INSERT INTO `tbl_hasilpenilaian` (`idpenilaian`, `nilaiakhir`, `keteranganhasil`) VALUES
(1, 21.00, 'TIDAK LULUS'),
(3, 331.00, 'lulus');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_kreteria`
--

CREATE TABLE IF NOT EXISTS `tbl_kreteria` (
  `idkreteria` varchar(2) NOT NULL,
  `namakreteria` varchar(20) NOT NULL,
  `skala` int(3) NOT NULL,
  `prioritas` varchar(5) NOT NULL,
  PRIMARY KEY (`idkreteria`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tbl_kreteria`
--

INSERT INTO `tbl_kreteria` (`idkreteria`, `namakreteria`, `skala`, `prioritas`) VALUES
('01', 'BAIK', 1, 'YA');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_pelamar`
--

CREATE TABLE IF NOT EXISTS `tbl_pelamar` (
  `kd_pelamar` int(30) NOT NULL,
  `nm_karyawan` varchar(50) NOT NULL,
  `ttl` varchar(150) NOT NULL,
  `alamat` varchar(100) NOT NULL,
  `telpon` varchar(14) NOT NULL,
  `jenis_kelamin` varchar(1) NOT NULL,
  `status` varchar(10) NOT NULL,
  `pendidikan` varchar(10) NOT NULL,
  `ipk` varchar(5) NOT NULL,
  `tgl_lamaran` date NOT NULL,
  PRIMARY KEY (`kd_pelamar`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tbl_pelamar`
--

INSERT INTO `tbl_pelamar` (`kd_pelamar`, `nm_karyawan`, `ttl`, `alamat`, `telpon`, `jenis_kelamin`, `status`, `pendidikan`, `ipk`, `tgl_lamaran`) VALUES
(11111, 'SURI', '111111', 'LEWORENG', '00000', 'W', 'QQ', 'D3', '3,0', '2011-11-01');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_penilaian`
--

CREATE TABLE IF NOT EXISTS `tbl_penilaian` (
  `kd_pelamar` varchar(5) NOT NULL,
  `tglpenilaian` date NOT NULL,
  `k1` varchar(5) NOT NULL,
  `k2` varchar(5) NOT NULL,
  `k3` varchar(5) NOT NULL,
  `k4` varchar(5) NOT NULL,
  `idpenilaian` int(11) NOT NULL,
  PRIMARY KEY (`kd_pelamar`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tbl_penilaian`
--

INSERT INTO `tbl_penilaian` (`kd_pelamar`, `tglpenilaian`, `k1`, `k2`, `k3`, `k4`, `idpenilaian`) VALUES
('11111', '2020-12-10', '98', '70', '74', '89', 1);

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_user`
--

CREATE TABLE IF NOT EXISTS `tbl_user` (
  `id` int(11) NOT NULL,
  `username` varchar(50) NOT NULL,
  `password` varchar(50) NOT NULL,
  `level` varchar(10) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Stand-in structure for view `view_hasil`
--
CREATE TABLE IF NOT EXISTS `view_hasil` (
`idpenilaian` int(11)
,`nm_karyawan` varchar(50)
,`k1` varchar(5)
,`k2` varchar(5)
,`k3` varchar(5)
,`k4` varchar(5)
,`nilaiakhir` decimal(5,2)
,`keteranganhasil` varchar(50)
);
-- --------------------------------------------------------

--
-- Stand-in structure for view `view_penilaian`
--
CREATE TABLE IF NOT EXISTS `view_penilaian` (
`kd_pelamar` int(30)
,`nm_karyawan` varchar(50)
,`tglpenilaian` date
,`k1` varchar(5)
,`k2` varchar(5)
,`k3` varchar(5)
,`k4` varchar(5)
);
-- --------------------------------------------------------

--
-- Struktur untuk view `view_hasil`
--
DROP TABLE IF EXISTS `view_hasil`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `view_hasil` AS select `tbl_penilaian`.`idpenilaian` AS `idpenilaian`,`tbl_pelamar`.`nm_karyawan` AS `nm_karyawan`,`tbl_penilaian`.`k1` AS `k1`,`tbl_penilaian`.`k2` AS `k2`,`tbl_penilaian`.`k3` AS `k3`,`tbl_penilaian`.`k4` AS `k4`,`tbl_hasilpenilaian`.`nilaiakhir` AS `nilaiakhir`,`tbl_hasilpenilaian`.`keteranganhasil` AS `keteranganhasil` from ((`tbl_hasilpenilaian` join `tbl_penilaian`) join `tbl_pelamar`);

-- --------------------------------------------------------

--
-- Struktur untuk view `view_penilaian`
--
DROP TABLE IF EXISTS `view_penilaian`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `view_penilaian` AS select `tbl_pelamar`.`kd_pelamar` AS `kd_pelamar`,`tbl_pelamar`.`nm_karyawan` AS `nm_karyawan`,`tbl_penilaian`.`tglpenilaian` AS `tglpenilaian`,`tbl_penilaian`.`k1` AS `k1`,`tbl_penilaian`.`k2` AS `k2`,`tbl_penilaian`.`k3` AS `k3`,`tbl_penilaian`.`k4` AS `k4` from (`tbl_pelamar` join `tbl_penilaian`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
