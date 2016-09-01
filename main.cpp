#include <gflags/gflags.h>
#include <libxl.h>
#include <boost/algorithm/string.hpp>
#include <boost/filesystem.hpp>
#include <fstream>
#include <iostream>
#include <string>
#include <vector>

DEFINE_string(d, "", "root directory");

using namespace std;

const char sep = ',';

struct Money {
    double shoukuan;
    double fahuo;
    double jieyu;
    Money() = default;
    ~Money() = default;
};

ostream& operator<<(ostream& os, const Money& money) {
    os << money.shoukuan << sep << money.fahuo << sep << money.jieyu;
    return os;
}

const char* format =
    "姓名,(货款)收款,(货款)发货,(货款)结余,(定位费)收款,"
    "(定位费)发货,(定位费)结余,(零售价任选)收款,(零售价任选)发货,("
    "零售价任选)结余,(运费)收款,(运费)发货,(运费)结余,(套盒奖励)收款,("
    "套盒奖励)发货,(套盒奖励)结余,店名,地址,电话号码\n";

class Agent {
   public:
    string name;
    string address;
    string shop;
    string phone;
    Money huokuan;
    Money dingweifei;
    Money lingshoujiarenxuan;
    Money yunfei;
    Money taohejiangli;

    libxl::Book* book;
    libxl::Sheet* sheet;
    string filename;

    Agent() = default;
    ~Agent() {
        sheet = nullptr;
        if (book != nullptr) {
            book->release();
            book = nullptr;
        }
    }

    Agent(Agent&& other) {
        name = std::move(other.name);
        address = std::move(other.address);
        shop = std::move(other.shop);
        phone = std::move(other.phone);
        huokuan = other.huokuan;
        dingweifei = other.dingweifei;
        lingshoujiarenxuan = other.lingshoujiarenxuan;
        yunfei = other.yunfei;
        taohejiangli = other.taohejiangli;
        filename = std::move(other.filename);
        this->book = other.book;
        other.book = nullptr;
        this->sheet = other.sheet;
        other.sheet = nullptr;
    }

    bool init(const string& fn) {
        filename = fn;
        book = xlCreateBook();
        if (book == nullptr) {
            return false;
        }
        if (!book->load(filename.c_str())) {
            return false;
        }
        sheet = book->getSheet(0);
        return true;
    }

    int parse() {
        try {
            if (sheet->cellType(1, 6) == libxl::CELLTYPE_STRING)
                name = sheet->readStr(1, 6, nullptr);
            if (sheet->cellType(1, 3) == libxl::CELLTYPE_STRING)
                shop = sheet->readStr(1, 3, nullptr);
            if (sheet->cellType(1, 12) == libxl::CELLTYPE_STRING)
                address = sheet->readStr(1, 12, nullptr);
            if (sheet->cellType(1, 9) == libxl::CELLTYPE_NUMBER) {
                uint64_t d_phone =
                    static_cast<uint64_t>(sheet->readNum(1, 9, nullptr));
                if (d_phone == 0) {
                    phone = "";
                } else {
                    phone = to_string(d_phone);
                }
            }
            huokuan.shoukuan = sheet->readNum(49, 4, nullptr);
            huokuan.fahuo = sheet->readNum(49, 5, nullptr);
            huokuan.jieyu = sheet->readNum(49, 6, nullptr);
            dingweifei.shoukuan = sheet->readNum(49, 7, nullptr);
            dingweifei.fahuo = sheet->readNum(49, 8, nullptr);
            dingweifei.jieyu = sheet->readNum(49, 9, nullptr);
            lingshoujiarenxuan.shoukuan = sheet->readNum(49, 10, nullptr);
            lingshoujiarenxuan.fahuo = sheet->readNum(49, 11, nullptr);
            lingshoujiarenxuan.jieyu = sheet->readNum(49, 12, nullptr);
            yunfei.shoukuan = sheet->readNum(49, 13, nullptr);
            yunfei.fahuo = sheet->readNum(49, 14, nullptr);
            yunfei.jieyu = sheet->readNum(49, 15, nullptr);
            taohejiangli.shoukuan = sheet->readNum(49, 16, nullptr);
            taohejiangli.fahuo = sheet->readNum(49, 17, nullptr);
            taohejiangli.jieyu = sheet->readNum(49, 18, nullptr);
        } catch (exception& e) {
            return -1;
        }
        return 0;
    }

    friend ostream& operator<<(ostream& os, const Agent& agent) {
        os << agent.name << agent.huokuan << sep << agent.dingweifei << sep
           << agent.lingshoujiarenxuan << sep << agent.yunfei << sep
           << agent.taohejiangli << sep << agent.shop << sep << agent.address
           << sep << agent.phone << '\n';
        return os;
    }
};

class Manager {
   public:
    string name;
    string district;
    vector<Agent> agents;

    Manager() = default;
    Manager(const string& _name, const string& _district)
        : name(_name), district(_district), agents() {}
    ~Manager() = default;

    void add_agent(const string& fn) {
        Agent agent;
        if (agent.init(fn) && agent.parse() == 0) {
            agents.push_back(std::move(agent));
        }
    }

    string build_ofn() const { return district + '_' + name + ".xls"; }

    int dump() {
        if (agents.size() == 0) {
            return 0;
        }
        string fn = std::move(build_ofn());
        libxl::Book* book = xlCreateBook();
        if (book == nullptr) {
            return -1;
        }
        libxl::Sheet* sheet = book->addSheet("Sheet1");
        if (sheet == nullptr) {
            book->release();
            return -1;
        }

        try {
            sheet->writeStr(0, 0, "feed the monster");
            sheet->writeStr(1, 0, "姓名");
            sheet->writeStr(1, 1, "(货款)收款");
            sheet->writeStr(1, 2, "(货款)发货");
            sheet->writeStr(1, 3, "(货款)结余");
            sheet->writeStr(1, 4, "(定位费)收款");
            sheet->writeStr(1, 5, "(定位费)发货");
            sheet->writeStr(1, 6, "(定位费)结余");
            sheet->writeStr(1, 7, "(零售价任选)收款");
            sheet->writeStr(1, 8, "(零售价任选)发货");
            sheet->writeStr(1, 9, "(零售价任选)结余");
            sheet->writeStr(1, 10, "(运费)收款");
            sheet->writeStr(1, 11, "(运费)发货");
            sheet->writeStr(1, 12, "(运费)结余");
            sheet->writeStr(1, 13, "(套盒奖励)收款");
            sheet->writeStr(1, 14, "(套盒奖励)发货");
            sheet->writeStr(1, 15, "(套盒奖励)结余");
            sheet->writeStr(1, 16, "店名");
            sheet->writeStr(1, 17, "地址");
            sheet->writeStr(1, 18, "电话号码");
            int row = 2;
            for (const Agent& agent : agents) {
                sheet->writeStr(row, 0, agent.name.c_str());
                sheet->writeNum(row, 1, agent.huokuan.shoukuan);
                sheet->writeNum(row, 2, agent.huokuan.fahuo);
                sheet->writeNum(row, 3, agent.huokuan.jieyu);
                sheet->writeNum(row, 4, agent.dingweifei.shoukuan);
                sheet->writeNum(row, 5, agent.dingweifei.fahuo);
                sheet->writeNum(row, 6, agent.dingweifei.jieyu);
                sheet->writeNum(row, 7, agent.lingshoujiarenxuan.shoukuan);
                sheet->writeNum(row, 8, agent.lingshoujiarenxuan.fahuo);
                sheet->writeNum(row, 9, agent.lingshoujiarenxuan.jieyu);
                sheet->writeNum(row, 10, agent.yunfei.shoukuan);
                sheet->writeNum(row, 11, agent.yunfei.fahuo);
                sheet->writeNum(row, 12, agent.yunfei.jieyu);
                sheet->writeNum(row, 13, agent.taohejiangli.shoukuan);
                sheet->writeNum(row, 14, agent.taohejiangli.fahuo);
                sheet->writeNum(row, 15, agent.taohejiangli.jieyu);
                sheet->writeStr(row, 16, agent.shop.c_str());
                sheet->writeStr(row, 17, agent.address.c_str());
                sheet->writeStr(row, 18, agent.phone.c_str());
                ++row;
            }
        } catch (exception& e) {
            book->release();
            return -1;
        }
        if (!book->save(build_ofn().c_str())) {
            book->release();
            return -1;
        }
        book->release();
        return 0;
    }

    friend ostream& operator<<(ostream& os, const Manager& manager) {
        os << "地区" << sep << "经理" << sep << format;
        for (const auto& agent : manager.agents) {
            os << manager.district << sep << manager.name << sep << agent;
        }
        return os;
    }
};

void walk_directory(const string& dir) {
    namespace bfs = boost::filesystem;
    bfs::path top(dir);
    bfs::directory_iterator iter(top), end;
    while (iter != end) {
        if (bfs::is_directory(iter->status())) {
            if (boost::algorithm::ends_with(iter->path().c_str(),
                                            ".DS_Store")) {
                continue;
            }
            string district = iter->path().leaf().c_str();
            Manager fake("", district);
            for (bfs::directory_iterator w(iter->path()), e; w != e; ++w) {
                if (bfs::is_directory(w->status()) &&
                    !boost::algorithm::ends_with(w->path().c_str(),
                                                 ".DS_Store")) {
                    string name = w->path().leaf().c_str();
                    Manager m(name, district);
                    for (bfs::directory_iterator ls(w->path()), le; ls != le;
                         ++ls) {
                        if (bfs::is_regular(ls->status()) &&
                            (boost::algorithm::ends_with(ls->path().c_str(),
                                                         ".xls") ||
                             boost::algorithm::ends_with(ls->path().c_str(),
                                                         ".xlsx"))) {
                            string fn(ls->path().c_str());
                            m.add_agent(fn);
                        }
                    }
                    m.dump();
                    // if (m.agents.size() > 0) {
                    // string ofn = std::move(m.build_ofn());
                    // fstream fs;
                    // fs.open(ofn, ios::out | ios::trunc);
                    // fs << m;
                    //}
                }

                if (bfs::is_regular(w->status()) &&
                    boost::algorithm::ends_with(w->path().c_str(), ".xls")) {
                    string fn(w->path().c_str());
                    fake.add_agent(fn);
                }
            }
            fake.dump();
            // if (fake.agents.size() > 0) {
            // string ofn = std::move(fake.build_ofn());
            // fstream fs;
            // fs.open(ofn, ios::out | ios::trunc);
            // fs << fake;
            //}
        }
        ++iter;
    }
}

int main(int argc, char** argv) {
    ::google::ParseCommandLineFlags(&argc, &argv, true);
    if (fLS::FLAGS_d.empty()) {
        cerr << "please enter root directory\n";
        exit(EXIT_FAILURE);
    }
    walk_directory(fLS::FLAGS_d);
    return 0;
}
