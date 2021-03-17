/*
1. Dosyalar olusturulur
2. önce write.js te datalar excele yazdirilir
3. read.js te exceldeki bilgiler okutulur.
*/

// Requiring module
const reader = require("xlsx");

// Reading our test file
const file = reader.readFile("./ornekDosya.xlsx");

// Sample data set
let student_data = [
  {
    ISIM: "Cabbar",
    SOYISIM: "Mikail",
    YAS: 22,
    "ALDIGI MAAS": 6000,
    CINSIYETI: "ERKEK",
  },
  {
    ISIM: "Hans",
    SOYISIM: "Joe",
    YAS: 39,
    "ALDIGI MAAS": 16000,
    CINSIYETI: "ERKEK",
  },
  {
    ISIM: "Murtaza",
    SOYISIM: "Kaya",
    YAS: 49,
    "ALDIGI MAAS": 6000,
    CINSIYETI: "ERKEK",
  },
  {
    ISIM: "Marion",
    SOYISIM: "Minna",
    YAS: 55,
    "ALDIGI MAAS": 9000,
    CINSIYETI: "KADIN",
  },
  {
    ISIM: "Murat",
    SOYISIM: "Burhan",
    YAS: 40,
    "ALDIGI MAAS": 10000,
    CINSIYETI: "ERKEK",
  },
  {
    ISIM: "Abdurrezzak",
    SOYISIM: "Adigüzel",
    YAS: 22,
    "ALDIGI MAAS": 6000,
    CINSIYETI: "ERKEK",
  },
  {
    ISIM: "Mehmet",
    SOYISIM: "Sökmen",
    YAS: 33,
    "ALDIGI MAAS": 12000,
    CINSIYETI: "ERKEK",
  },
];

const ws = reader.utils.json_to_sheet(student_data);

reader.utils.book_append_sheet(file, ws, "Sheet3");

// Writing to our file
reader.writeFile(file, "./ornekDosya.xlsx");
