using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lab7
{
    public partial class Form1 : MaterialForm
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\nikli\Desktop\Лабы\lab_3.accdb;Jet OLEDB:Create System Database=true;Jet OLEDB:System database=C:\Users\nikli\AppData\Roaming\Microsoft\Access\System.mdw");
        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet ds;
        bool added = true;
        bool updated = true;
        bool deleted = true;

        public Form1()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue400, Primary.Blue400, Accent.LightBlue200, TextShade.WHITE);
        }
        void GetAllLists()
        {
            da = new OleDbDataAdapter("SELECT * FROM Доставка", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "Доставка");

            dataGridView1.DataSource = ds.Tables["Доставка"].DefaultView;
            da = new OleDbDataAdapter("SELECT * FROM Покупатель", con);
            ds = new DataSet();
            da.Fill(ds, "Покупатель");
            dataGridView2.DataSource = ds.Tables["Покупатель"].DefaultView;

            da = new OleDbDataAdapter("SELECT * FROM Продавец", con);
            ds = new DataSet();
            da.Fill(ds, "Продавец");
            dataGridView3.DataSource = ds.Tables["Продавец"].DefaultView;

            da = new OleDbDataAdapter("SELECT * FROM Продажа", con);
            ds = new DataSet();
            da.Fill(ds, "Продажа");
            dataGridView4.DataSource = ds.Tables["Продажа"].DefaultView;

            da = new OleDbDataAdapter("SELECT * FROM Склад", con);
            ds = new DataSet();
            da.Fill(ds, "Склад");
            dataGridView5.DataSource = ds.Tables["Склад"].DefaultView;

            da = new OleDbDataAdapter("SELECT * FROM Товар", con);
            ds = new DataSet();
            da.Fill(ds, "Товар");
            dataGridView6.DataSource = ds.Tables["Товар"].DefaultView;

            try
            {
                da = new OleDbDataAdapter("SELECT Name AS 'Запросы' FROM MSysObjects WHERE Type = 5", con);
                ds = new DataSet();
                da.Fill(ds, "[MSysObjects]");
                dataGridView7.DataSource = ds.Tables[0].DefaultView;
            }
            catch
            {

            }
            con.Close();      
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            GetAllLists();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";
        }

        //INSERT Доставка
        private void materialFlatButton1_Click(object sender, EventArgs e)
        {
            if(textBox3.Text == "")
            {
                con.Open();
                try
                {
                    added = true;
                    materialLabel4.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO Доставка ( ПунктОтправления, ПунктНазначения, Дата, ОтметкаОвыполнении )" +
                        "VALUES(@from, @where, DateValue(@date), @done)";
                    cmd.Parameters.AddWithValue("@from", textBox2.Text);
                    cmd.Parameters.AddWithValue("@where", textBox1.Text);
                    cmd.Parameters.AddWithValue("@date", dateTimePicker1.Text);
                    if (materialCheckBox1.CheckState == CheckState.Checked)
                    {
                        cmd.Parameters.AddWithValue("@done", "выполнено");
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@done", "выполняется");
                    }
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    added = false;
                    materialLabel4.Text = "Ошибка при добавлении";
                }
                finally
                {
                    if(added)
                    {
                        materialLabel4.Text = "Добавлено";
                    }
                    con.Close();
                }
            }
        }

        //UPDATE Доставка
        private void materialFlatButton2_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                updated = true;
                con.Open();
                try
                {
                    materialLabel4.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "UPDATE Доставка SET ПунктОтправления = @from, ПунктНазначения = @where, Дата = DateValue(@date), ОтметкаОвыполнении = @done WHERE КодДоставки = @end";         
                    if(textBox2.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@from", textBox2.Text);
                    }
                    if(textBox1.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@where", textBox1.Text);
                    }
                    if(dateTimePicker1.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@date", dateTimePicker1.Text);
                    }
                    if (materialCheckBox1.CheckState == CheckState.Checked)
                    {
                        cmd.Parameters.AddWithValue("@done", "выполнено");
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@done", "выполняется");
                    }
                    cmd.Parameters.AddWithValue("@end", textBox3.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    updated = false;
                    materialLabel4.Text = "Ошибка при изменении";
                }
                finally
                {
                    if(updated)
                    {
                        materialLabel4.Text = "Изменено";
                    }
                    con.Close();
                }
            }
        }

        //DELETE Доставка
        private void materialFlatButton3_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                deleted = true;
                con.Open();
                try
                {
                    materialLabel4.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "DELETE FROM Доставка WHERE КодДоставки = " + textBox3.Text;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    deleted = false;
                    materialLabel4.Text = "Ошибка при удалении";
                }
                finally
                {
                    if(deleted)
                    {
                        materialLabel4.Text = "Удалено";
                    }
                    con.Close();
                }
            }
        }

        //INSERT Покупатель
        private void materialFlatButton4_Click(object sender, EventArgs e)
        {
            if (textBox8.Text == "")
            {
                added = true;
                con.Open();
                try
                {
                    materialLabel10.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO Покупатель ( НаименованиеПокупателя, Адрес, Телефон, ИНН )" +
                        "VALUES(@name, @address, @phone, @INN)";
                    cmd.Parameters.AddWithValue("@name", textBox7.Text);
                    cmd.Parameters.AddWithValue("@address", textBox6.Text);
                    cmd.Parameters.AddWithValue("@phone", textBox5.Text);
                    cmd.Parameters.AddWithValue("@INN", textBox4.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    added = false;
                    materialLabel10.Text = "Ошибка при добавлении";
                }
                finally
                {
                    if (added)
                    {
                        materialLabel10.Text = "Добавлено";
                    }
                    con.Close();
                }
            }
        }

        //UPDATE Покупатель
        private void materialFlatButton5_Click(object sender, EventArgs e)
        {
            if (textBox8.Text != "")
            {
                updated = true;
                con.Open();
                try
                {
                    materialLabel10.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "UPDATE Покупатель SET НаименованиеПокупателя = @name, Адрес = @address, Телефон = @phone, ИНН = @INN WHERE КодПокупателя = @end";
                    if (textBox7.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@name",  textBox7.Text);
                    }
                    if (textBox6.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@address", textBox6.Text);
                    }
                    if (textBox5.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@phone", textBox5.Text);
                    }
                    if (textBox4.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@INN", textBox4.Text);
                    }
                    cmd.Parameters.AddWithValue("@end", textBox8.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    updated = false;
                    materialLabel10.Text = "Ошибка при изменении";
                }
                finally
                {
                    if (updated)
                    {
                        materialLabel10.Text = "Изменено";
                    }
                    con.Close();
                }
            }
        }

        //DELETE Покупатель
        private void materialFlatButton6_Click(object sender, EventArgs e)
        {
            if (textBox8.Text != "")
            {
                deleted = true;
                con.Open();
                try
                {
                    materialLabel10.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "DELETE FROM Покупатель WHERE КодПокупателя = " + textBox8.Text;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    deleted = false;
                    materialLabel10.Text = "Ошибка при удалении";
                }
                finally
                {
                    if (deleted)
                    {
                        materialLabel10.Text = "Удалено";
                    }
                    con.Close();
                }
            }
        }

        //INSERT Продавец
        private void materialFlatButton7_Click(object sender, EventArgs e)
        {
            if (textBox9.Text == "")
            {
                added = true;
                con.Open();
                try
                {
                    materialLabel11.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO Продавец ( НаименованиеПродавца, Адрес, Телефон, ИНН )" +
                        "VALUES(@name, @address, @phone, @INN)";
                    cmd.Parameters.AddWithValue("@name", textBox10.Text);
                    cmd.Parameters.AddWithValue("@address", textBox11.Text);
                    cmd.Parameters.AddWithValue("@phone", textBox12.Text);
                    cmd.Parameters.AddWithValue("@INN", textBox13.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    added = false;
                    materialLabel11.Text = "Ошибка при добавлении";
                }
                finally
                {
                    if (added)
                    {
                        materialLabel11.Text = "Добавлено";
                    }
                    con.Close();
                }
            }
        }

        //UPDATE Продавец
        private void materialFlatButton8_Click(object sender, EventArgs e)
        {
            if (textBox9.Text != "")
            {
                updated = true;
                con.Open();
                try
                {
                    materialLabel11.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "UPDATE Продавец SET НаименованиеПродавца = @name, Адрес = @address, Телефон = @phone, ИНН = @INN WHERE КодПродавца = @end";
                    if (textBox10.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@name", textBox10.Text);
                    }
                    if (textBox11.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@address", textBox11.Text);
                    }
                    if (textBox12.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@phone", textBox12.Text);
                    }
                    if (textBox13.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@INN", textBox13.Text);
                    }
                    cmd.Parameters.AddWithValue("@end", textBox9.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    updated = false;
                    materialLabel11.Text = "Ошибка при изменении";
                }
                finally
                {
                    if(updated)
                    {
                        materialLabel11.Text = "Изменено";
                    }
                    con.Close();
                }
            }
        }

        //DELETE Продавец
        private void materialFlatButton9_Click(object sender, EventArgs e)
        {
            if (textBox9.Text != "")
            {
                deleted = true;
                con.Open();
                try
                {
                    materialLabel11.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "DELETE FROM Продавец WHERE КодПродавца = " + textBox9.Text;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    deleted = false;
                    materialLabel11.Text = "Ошибка при удалении";
                }
                finally
                {
                    if (deleted)
                    {
                        materialLabel11.Text = "Удалено";
                    }
                    con.Close();
                }
            }
        }

        //INSERT Продажа
        private void materialFlatButton10_Click(object sender, EventArgs e)
        {
            if (textBox14.Text == "")
            {
                added = true;
                con.Open();
                try
                {
                    materialLabel18.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO Продажа (КодТовара, КодПродавца, КодПокупателя, КодДоставки, Цена, Количество, Дата )" +
                        "VALUES(@KT, @KPr, @KPo, @KD, @price, @count, DateValue(@date))";
                    cmd.Parameters.AddWithValue("@KT", textBox15.Text);
                    cmd.Parameters.AddWithValue("@KPr", textBox16.Text);
                    cmd.Parameters.AddWithValue("@KPo", textBox17.Text);
                    cmd.Parameters.AddWithValue("@KD", textBox18.Text);
                    cmd.Parameters.AddWithValue("@price", textBox19.Text);
                    cmd.Parameters.AddWithValue("@count", textBox20.Text);
                    cmd.Parameters.AddWithValue("@date", dateTimePicker2.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    added = false;
                    materialLabel18.Text = "Ошибка при добавлении";
                }
                finally
                {
                    if (added)
                    {
                        materialLabel18.Text = "Добавлено";
                    }
                    con.Close();
                }
            }
        }

        //UPDATE Продажа
        private void materialFlatButton11_Click(object sender, EventArgs e)
        {
            if (textBox14.Text != "")
            {
                updated = true;
                con.Open();
                try
                {
                    materialLabel18.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "UPDATE Продажа SET КодТовара = @KT, КодПродавца = @KPr, КодПокупателя = @KPo, КодДоставки = @KD, Цена = @price, Количество = @count, Дата = DateValue(@date) WHERE НомерНакладной = @end";
                    if (textBox15.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@KT", textBox15.Text);
                    }
                    if (textBox16.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@KPr", textBox16.Text);
                    }
                    if (textBox17.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@KPo", textBox17.Text);
                    }
                    if (textBox18.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@KD", int.Parse(textBox18.Text));
                    }
                    if (textBox19.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@price", int.Parse(textBox19.Text));
                    }
                    if (textBox20.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@count", int.Parse(textBox20.Text));
                    }
                    cmd.Parameters.AddWithValue("@end", textBox14.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    updated = false;
                    materialLabel18.Text = "Ошибка при изменении";
                }
                finally
                {
                    if(updated)
                    {
                        materialLabel18.Text = "Изменено";
                    }
                    con.Close();
                }
            }
        }

        //DELETE Продажа
        private void materialFlatButton12_Click(object sender, EventArgs e)
        {
            if (textBox14.Text != "")
            {
                deleted = true;
                con.Open();
                try
                {
                    materialLabel18.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "DELETE FROM Продажа WHERE НомерНакладной = " + textBox14.Text;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    deleted = false;
                    materialLabel18.Text = "Ошибка при удалении";
                }
                finally
                {
                    if(deleted)
                    {
                        materialLabel18.Text = "Удалено";
                    }
                    con.Close();
                }
            }
        }

        //INSERT Склад
        private void materialFlatButton13_Click(object sender, EventArgs e)
        {
            if (textBox21.Text == "")
            {
                added = true;
                con.Open();
                try
                {
                    materialLabel35.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO Склад ( Регион, Адрес, Телефон, Площадь )" +
                        "VALUES(@region, @address, @phone, @square)";
                    cmd.Parameters.AddWithValue("@region", textBox22.Text);
                    cmd.Parameters.AddWithValue("@address", textBox23.Text);
                    cmd.Parameters.AddWithValue("@phone", textBox24.Text);
                    cmd.Parameters.AddWithValue("@square", textBox25.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    added = false;
                    materialLabel35.Text = "Ошибка при добавлении";
                }
                finally
                {
                    if(added)
                    {
                        materialLabel35.Text = "Добавлено";
                    }
                    con.Close();
                }
            }
        }

        //UPDATE Склад
        private void materialFlatButton14_Click(object sender, EventArgs e)
        {
            if (textBox21.Text != "")
            {
                updated = true;
                con.Open();
                try
                {
                    materialLabel35.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "UPDATE Склад SET Регион = @region, Адрес = @address, Телефон = @phone, Площадь = @square WHERE НомерСклада = @end";
                    if (textBox22.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@region", textBox22.Text);
                    }
                    if (textBox23.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@address", textBox23.Text);
                    }
                    if (textBox24.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@phone", textBox24.Text);
                    }
                    if (textBox25.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@square", textBox25.Text);
                    }
                    cmd.Parameters.AddWithValue("@end", textBox21.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    updated = false;
                    materialLabel35.Text = "Ошибка при изменении";
                }
                finally
                {
                    if(updated)
                    {
                        materialLabel35.Text = "Изменено";
                    }
                    con.Close();
                }
            }
        }

        //DELETE Склад
        private void materialFlatButton15_Click(object sender, EventArgs e)
        {
            if (textBox21.Text != "")
            {
                deleted = true;
                con.Open();
                try
                {
                    materialLabel35.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "DELETE FROM Склад WHERE НомерСклада = " + textBox21.Text;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    deleted = false;
                    materialLabel35.Text = "Ошибка при удалении";
                }
                finally
                {
                    if(deleted)
                    {
                        materialLabel35.Text = "Удалено";
                    }
                    con.Close();
                }
            }
        }

        //INSERT Товар
        private void materialFlatButton16_Click(object sender, EventArgs e)
        {
            if (textBox26.Text == "")
            {
                added = true;
                con.Open();
                try
                {
                    materialLabel36.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "INSERT INTO Товар ( НаименованиеТовара, Вес, Размеры, НомерСклада )" +
                        "VALUES(@name, @mass, @size, @warehouse)";
                    cmd.Parameters.AddWithValue("@name", textBox27.Text);
                    cmd.Parameters.AddWithValue("@mass", int.Parse(textBox28.Text));
                    cmd.Parameters.AddWithValue("@size", textBox29.Text);
                    cmd.Parameters.AddWithValue("@warehouse", textBox30.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    added = false;
                    materialLabel36.Text = "Ошибка при добавлении";
                }
                finally
                {
                    if(added)
                    {
                        materialLabel36.Text = "Добавлено";
                    }
                    con.Close();
                }
            }
        }

        //UPDATE Товар
        private void materialFlatButton17_Click(object sender, EventArgs e)
        {
            if (textBox26.Text != "")
            {
                updated = true;
                con.Open();
                try
                {
                    materialLabel36.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "UPDATE Товар SET НаименованиеТовара = @name, Вес = @mass, Размеры = @size, НомерСклада = @warehouse WHERE КодТовара = @end";
                    if (textBox27.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@name", textBox27.Text);
                    }
                    if (textBox28.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@mass", int.Parse(textBox28.Text));
                    }
                    if (textBox29.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@size", textBox29.Text);
                    }
                    if (textBox30.Text != "")
                    {
                        cmd.Parameters.AddWithValue("@warehouse", textBox30.Text);
                    }
                    cmd.Parameters.AddWithValue("@end", textBox26.Text);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    updated = false;
                    materialLabel36.Text = "Ошибка при изменении";
                }
                finally
                {
                    if(updated)
                    {
                        materialLabel36.Text = "Изменено";
                    }
                    con.Close();
                }
            }
        }

        //DELETE Товар
        private void materialFlatButton18_Click(object sender, EventArgs e)
        {
            if (textBox26.Text != "")
            {
                deleted = true;
                con.Open();
                try
                {
                    materialLabel36.Text = "";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "DELETE FROM Товар WHERE КодТовара = " + textBox26.Text;
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    deleted = false;
                    materialLabel36.Text = "Ошибка при удалении";
                }
                finally
                {
                    if (deleted)
                    {
                        materialLabel36.Text = "Удалено";
                    }             
                    con.Close();
                }
            }
        }

        //Процедура
        private void materialFlatButton19_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "EXEC " + textBox31.Text;
                cmd.ExecuteNonQuery();
            }
            catch(Exception exp)
            {
                MessageBox.Show(exp.ToString());
            }
            finally
            {
                con.Close();
            }
        }

        private void materialRaisedButton1_Click(object sender, EventArgs e)
        {
            GetAllLists();
        }

        //Функция
        private void materialFlatButton20_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                cmd.Parameters.AddWithValue("x1", int.Parse(textBox32.Text));
                cmd.CommandText = "SELECT * FROM " + textBox31.Text + " AS Запрос";
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds, "[Запрос]");
                dataGridView7.DataSource = ds.Tables[0].DefaultView;
            }
            catch(Exception exp)
            {
                MessageBox.Show(exp.ToString());
            }
        }

        //Представление
        private void materialFlatButton21_Click(object sender, EventArgs e)
        {
            try
            {
                string query = "SELECT * FROM " + textBox31.Text + " AS Запрос";
                OleDbCommand cmd = new OleDbCommand(query, con);
                OleDbDataAdapter da = new OleDbDataAdapter(query, con);
                DataSet ds = new DataSet();
                da.Fill(ds, "[Запрос]");
                dataGridView7.DataSource = ds.Tables[0].DefaultView;
            }
            catch(Exception exp)
            {
                MessageBox.Show(exp.ToString());
            }
        }
    }
}
