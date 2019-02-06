use calamine::{open_workbook_auto, Reader};
use regex::Regex;
use std::path::{Path, PathBuf};
use structopt::StructOpt;

#[derive(StructOpt, Debug)]
#[structopt(name = "Spreadsheet Finder")]
struct Opt {
    /// Regex for cell value to find
    #[structopt(short = "r", long = "regex")]
    regex: Regex,

    /// Set file matcher regex
    #[structopt(short = "n", long = "fileregex", default_value = r".*\.ods")]
    filematch: Regex,

    /// Folder to process
    #[structopt(name = "FOLDER", parse(from_os_str))]
    folder: PathBuf,
}

fn main() {
    let opt = Opt::from_args();
    let regex = opt.regex;
    let file_regex = opt.filematch;
    let mut start_path = opt.folder;
    find_children(&mut start_path, &|v: &Path| {
        if let Some(name) = v.file_name() {
            if file_regex.is_match(&name.to_string_lossy()) {
                read(v, &regex);
            }
        }
    });
}

fn find_children<F>(path: &mut PathBuf, callback: &F)
where
    F: Fn(&Path),
{
    if path.is_dir() {
        for f in path.read_dir().expect("Can't read dir!") {
            let mut f = f.expect("Can't read dir entry!").path();
            if f.is_dir() {
                find_children(&mut f, callback);
            } else if f.is_file() {
                callback(&f);
            } else {
                eprint!("Ignoring {}", f.to_string_lossy());
            }
        }
    }
}

fn read(path: &Path, regex: &Regex) {
    let mut document = match open_workbook_auto(path) {
        Ok(v) => v,
        Err(e) => {
            eprintln!("Unable to process file {}: {}", path.to_string_lossy(), e);
            return;
        }
    };
    let amount = document.sheet_names().len();
    for i in 0..amount {
        let sheet = document.sheet_names()[i].clone();
        match document.worksheet_range(&sheet) {
            Some(Ok(r)) => {
                for row in r.rows() {
                    for cell in row {
                        let value = format!("{}", cell);
                        if regex.is_match(&value) {
                            println!("Match for {}:\t {}", path.to_string_lossy(), cell);
                        }
                    }
                }
            }
            None => eprintln!("No reader for {}", path.to_string_lossy()),
            Some(Err(e)) => eprintln!("Error reading file {}: {}", path.to_string_lossy(), e),
        }
    }
}
