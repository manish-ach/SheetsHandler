use calamine::{open_workbook_auto, Reader};
use colored::Colorize;
use csv::ReaderBuilder;
use lettre::{transport::smtp::authentication::Credentials, Message, SmtpTransport, Transport};
use std::error::Error;
use std::{
    io::{self, Write},
    process::exit,
};
use term_size;

fn main() {
    clear_screen();
    let arg: Vec<String> = std::env::args().collect();

    if arg.len() < 2 {
        indv();
    } else if arg[1] == "-a" {
        clear_screen();
        all();
        exit(0);
    }
}

fn all() {
    let path = "/Users/devadmin/software/details.xlsx";
    let mut workbook = open_workbook_auto(path).unwrap();
    if let Some(Ok(sheet)) = workbook.worksheet_range_at(0) {
        println!("{}", "All Records:".green());
        for row in sheet.rows() {
            let roll_no = row.get(0).map(|cell| cell.to_string()).unwrap_or_default();
            let name = row.get(1).map(|cell| cell.to_string()).unwrap_or_default();
            let email = row.get(2).map(|cell| cell.to_string()).unwrap_or_default();
            let alternate_email = row.get(3).map(|cell| cell.to_string()).unwrap_or_default();
            let password = row.get(4).map(|cell| cell.to_string()).unwrap_or_default();

            print!(
                "{}  {}  {}  {}  {}",
                roll_no.red(),
                name.blue(),
                email.red(),
                alternate_email.green(),
                password.white()
            );

            println!();
        }
    } else {
        println!("Failed to read the worksheet.");
    }
}
fn scn() -> String {
    println!("\n\t\t{}", "Enter Roll no or Name".green().bold());
    print!("\t\t{}", "=> ".blue());
    io::stdout().flush().expect("error");
    let mut search = String::new();
    io::stdin().read_line(&mut search).expect("eror");
    search = search.trim().to_lowercase();
    search
}
fn lops() -> bool {
    let mut lop = true;
    println!("\n\t\t{}", "Want to continue? (y/n)".green());
    print!("\t\t{}", "=> ".blue());
    io::stdout().flush().expect("err");
    let mut ch = String::new();
    io::stdin().read_line(&mut ch).expect("failure");

    if ch.trim().to_lowercase() != "y" {
        lop = false;
    }
    lop
}

fn indv() {
    let mut lop = true;
    while lop {
        clear_screen();

        let mut roll_no = String::from("-");
        let mut name = String::from("-");
        let mut email = String::from("-");
        let mut alternate_email = String::from("-");
        let mut password = String::from("-");

        top_screen();
        data_screen(&roll_no, &name, &email, &alternate_email, &password);
        let search = scn();

        let path = "/Users/devadmin/software/details.xlsx";
        if let Ok(mut workbook) = open_workbook_auto(path) {
            if let Some(Ok(sheet)) = workbook.worksheet_range_at(0) {
                let mut found = false;

                for row in sheet.rows() {
                    if row.len() >= 2 {
                        let col1 = row[0].to_string().trim().to_lowercase();
                        let col2 = row[1].to_string().trim().to_lowercase();

                        if col1.contains(&search) || col2.contains(&search) {
                            roll_no = row.get(0).map(|cell| cell.to_string()).unwrap_or_default();
                            name = row.get(1).map(|cell| cell.to_string()).unwrap_or_default();
                            email = row.get(2).map(|cell| cell.to_string()).unwrap_or_default();
                            alternate_email =
                                row.get(3).map(|cell| cell.to_string()).unwrap_or_default();
                            password = row.get(4).map(|cell| cell.to_string()).unwrap_or_default();

                            found = true;
                            break;
                        }
                    }
                }

                clear_screen();
                top_screen();
                if found {
                    data_screen(&roll_no, &name, &email, &alternate_email, &password);
                } else {
                    println!("\n\t{}", "No matching records found.".red());
                }
            }
        }

        lop = lops();
    }
}
fn top_screen() {
    let header = String::from("***** Student Management System *****");
    let spaces = space(&header);
    println!("\n\n{}{}\n\n", spaces, header.black().bold().on_cyan());
}
fn data_screen(roll_no: &str, name: &str, email: &str, alternate_email: &str, password: &str) {
    let strn = String::from("Current Data:");
    let spaces = space(&strn);
    println!("{}{}", spaces, strn.on_blue().bold().bright_red());
    println!("\t\t{}{}", "roll no: ".green(), roll_no.red());
    println!("\t\t{}{}", "name: ".green(), name.red());
    println!("\t\t{}{}", "email: ".green(), email.red());
    println!(
        "\t\t{}{}",
        "alternate email: ".green(),
        alternate_email.red()
    );
    println!("\t\t{}{}", "password: ".green(), password.red());
}
fn clear_screen() {
    std::process::Command::new("clear")
        .status()
        .expect("error clearing screen");
}
fn space(text: &str) -> String {
    if let Some((width, _)) = term_size::dimensions() {
        let text_len = text.len();
        let spaces = (width - text_len) / 2;

        let padding: String = " ".repeat(spaces);
        return padding;
    } else {
        return String::from("error");
    }
}

// use std::io::Write;
// use std::{env, fs, path, process};
// //in main
// if arg[0] != "sms" && false {
//     process::exit(0);
// }
// let path = path::Path::new("/Users/devadmin/software");
// if !path.is_dir() {
//     fs::create_dir_all(path).expect("creating dir failed");
// }
// if arg[1] == "-n" || true {
//     let name = arg[2].to_string();
//     add_record(name);
// }

// fn students() {}
// fn add_record(name: String) {
//     let path = path::Path::new("/Users/devadmin/software/name.txt");
//     if !path.exists() {
//         fs::File::create(path).expect("error creating file");
//     }
//     let mut file = fs::OpenOptions::new()
//         .append(true)
//         .create(true)
//         .open(path)
//         .expect("opening file error");
//     writeln!(file, "{}", name).expect("eror");
// }
