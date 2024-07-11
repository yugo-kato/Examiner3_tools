use std::fs::File;
use std::io::{BufRead, BufReader, Write};
use calamine::{DataType, Reader, open_workbook, Xlsx};
use std::io;
use std::path::PathBuf;
use rfd::{FileDialog};

fn csv_cut(openfilename: &str, index_x: i32, index_x2: i32) {
    let outputfilename = format!("{}_X{}-{}.csv", &openfilename[..openfilename.len()-4], index_x, index_x2);

    let f = File::open(openfilename).expect("Failed to open input file");
    let mut out = File::create(&outputfilename).expect("Failed to create output file");

    let reader = BufReader::new(f);
    let mut lines = reader.lines();

    if let Some(Ok(line)) = lines.next() {
        writeln!(out, "{}", line).expect("Failed to write to output file");
    }

    for line in lines {
        if let Ok(line) = line {
            let tmp: Vec<&str> = line.split(',').collect();
            if let (Ok(x), Ok(z)) = (tmp[0].parse::<i32>(), tmp[2].parse::<i32>()) {
                if index_x <= x && x <= index_x2 && z <= 93 {
                    writeln!(out, "{}", line).expect("Failed to write to output file");
                }
            }
        }
    }
}

fn main() -> io::Result<()> {

    let input_file_path = match FileDialog::new().add_filter("Excel file",&["xlsx"]).pick_file() {
        Some(path) => path,
        None => return Err(io::Error::new(io::ErrorKind::Other, "No file selected.")),
    };


    let mut df: Vec<(String, i32, i32)> = vec![];

    // Load the Excel file
    let mut workbook: Xlsx<_> = open_workbook(input_file_path.clone()).expect("Failed to open Excel file");

    // Select the first sheet and read the data
    if let Some(Ok(r)) = workbook.worksheet_range_at(0) {
        for row in r.rows().skip(1) {
            let name = row[0].get_string().unwrap().to_string();
            let start_x = row[2].get_float().unwrap() as i32;
            let end_x = row[3].get_float().unwrap() as i32;
            df.push((name, start_x, end_x));
        }
    }

    for (i, (name, start_x, end_x)) in df.iter().enumerate() {
        println!("{}: {}, {}-{}", i, name, start_x, end_x);
    
        let mut str_path = PathBuf::from(input_file_path.parent().unwrap());
        str_path.push(format!(r"{}.csv", name));

        //let str_path = input_file_path.parent().unwrap().to_str().unwrap().to_owned() + format!(r"\{}.csv", name);
        csv_cut(&str_path.to_str().unwrap(), *start_x, *end_x);
    }

    Ok(())

}

