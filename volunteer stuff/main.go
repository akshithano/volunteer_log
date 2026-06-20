package main

import (
	"bytes"
	"database/sql"
	"fmt"
	"html/template"
	"io"
	"log"
	"net/http"
	"net/smtp"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/joho/godotenv"
	"github.com/tealeg/xlsx"
	_ "github.com/lib/pq"
)

var db *sql.DB

func getEnv(key, fallback string) string {
	if v := os.Getenv(key); v != "" {
		return v
	}
	return fallback
}

var (
	smtpHost = getEnv("SMTP_HOST", "smtp.gmail.com")
	smtpPort = getEnv("SMTP_PORT", "587")
	smtpUser = getEnv("SMTP_USER", "")
	smtpPass = getEnv("SMTP_PASS", "")
	smtpFrom = getEnv("SMTP_FROM", "Thuvakkam Volunteers <noreply@thuvakkam.org>")
)

type Volunteer struct {
	ID         int
	Name       string
	Email      string
	Mobile     string
	Hours      int
	ProfilePic string
}

type PageData struct {
	Volunteer Volunteer
	Section   string
	Flash     string
}

type AdminPageData struct {
	Volunteers []Volunteer
	Flash      string
}

func getSessionEmail(r *http.Request) string {
	c, err := r.Cookie("vol_email")
	if err != nil {
		return ""
	}
	return c.Value
}

func setSession(w http.ResponseWriter, email string) {
	http.SetCookie(w, &http.Cookie{
		Name: "vol_email", Value: email,
		Path: "/", HttpOnly: true, MaxAge: 86400 * 7,
	})
}

func clearSession(w http.ResponseWriter) {
	http.SetCookie(w, &http.Cookie{Name: "vol_email", Value: "", Path: "/", MaxAge: -1})
}

func getVolunteer(r *http.Request) (Volunteer, error) {
	email := getSessionEmail(r)
	if email == "" {
		return Volunteer{}, fmt.Errorf("not logged in")
	}
	var v Volunteer
	err := db.QueryRow(
		"SELECT volunteer_id, name, email, COALESCE(mobile,''), total_hours, COALESCE(profile_pic,'') FROM volunteers WHERE email=$1",
		email,
	).Scan(&v.ID, &v.Name, &v.Email, &v.Mobile, &v.Hours, &v.ProfilePic)
	return v, err
}

func getAllVolunteers() ([]Volunteer, error) {
	rows, err := db.Query("SELECT volunteer_id, name, email, total_hours FROM volunteers ORDER BY total_hours DESC")
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	var list []Volunteer
	for rows.Next() {
		var v Volunteer
		rows.Scan(&v.ID, &v.Name, &v.Email, &v.Hours)
		list = append(list, v)
	}
	return list, nil
}

func isAdmin(email string) bool {
	raw := getEnv("ADMIN_EMAILS", "")
	if raw == "" {
		return false
	}
	for _, e := range strings.Split(raw, ",") {
		if strings.TrimSpace(e) == email {
			return true
		}
	}
	return false
}

func buildExcel(volunteers []Volunteer) (*xlsx.File, error) {
	f := xlsx.NewFile()
	sheet, err := f.AddSheet("Volunteers")
	if err != nil {
		return nil, err
	}
	header := sheet.AddRow()
	for _, h := range []string{"ID", "Name", "Email", "Total Hours"} {
		cell := header.AddCell()
		cell.Value = h
		style := cell.GetStyle()
		style.Font.Bold = true
		cell.SetStyle(style)
	}
	for _, v := range volunteers {
		row := sheet.AddRow()
		row.AddCell().SetInt(v.ID)
		row.AddCell().Value = v.Name
		row.AddCell().Value = v.Email
		row.AddCell().SetInt(v.Hours)
	}
	return f, nil
}

func sendExcelEmail() {
	volunteers, err := getAllVolunteers()
	if err != nil {
		log.Println("Weekly email: DB error:", err)
		return
	}
	f, err := buildExcel(volunteers)
	if err != nil {
		log.Println("Weekly email: excel build error:", err)
		return
	}
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		log.Println("Weekly email: excel write error:", err)
		return
	}
	xlsxBytes := buf.Bytes()

	boundary := "ThuvakkamBoundary42"
	subject := fmt.Sprintf("Weekly Volunteer Report — %s", time.Now().Format("02 Jan 2006"))

	var msg bytes.Buffer
	msg.WriteString("From: " + smtpFrom + "\r\n")
	msg.WriteString("To: " + getEnv("ADMIN_EMAILS", smtpUser) + "\r\n")
	msg.WriteString("Subject: " + subject + "\r\n")
	msg.WriteString("MIME-Version: 1.0\r\n")
	msg.WriteString(fmt.Sprintf("Content-Type: multipart/mixed; boundary=\"%s\"\r\n\r\n", boundary))
	msg.WriteString(fmt.Sprintf("--%s\r\n", boundary))
	msg.WriteString("Content-Type: text/plain; charset=UTF-8\r\n\r\n")
	msg.WriteString(fmt.Sprintf("Hi,\n\nPlease find this week's volunteer hours report attached.\n\nTotal volunteers: %d\n\nRegards,\nThuvakkam Dashboard", len(volunteers)))
	msg.WriteString("\r\n")
	msg.WriteString(fmt.Sprintf("--%s\r\n", boundary))
	msg.WriteString("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n")
	msg.WriteString("Content-Transfer-Encoding: base64\r\n")
	msg.WriteString(fmt.Sprintf("Content-Disposition: attachment; filename=\"volunteers_%s.xlsx\"\r\n\r\n", time.Now().Format("2006-01-02")))

	b64encode := func(data []byte) string {
		const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
		var result strings.Builder
		for i := 0; i < len(data); i += 3 {
			var b [3]byte
			n := copy(b[:], data[i:])
			result.WriteByte(chars[b[0]>>2])
			result.WriteByte(chars[(b[0]&0x3)<<4|b[1]>>4])
			if n > 1 {
				result.WriteByte(chars[(b[1]&0xF)<<2|b[2]>>6])
			} else {
				result.WriteByte('=')
			}
			if n > 2 {
				result.WriteByte(chars[b[2]&0x3F])
			} else {
				result.WriteByte('=')
			}
			if (i/3+1)%19 == 0 {
				result.WriteString("\r\n")
			}
		}
		return result.String()
	}
	msg.WriteString(b64encode(xlsxBytes))
	msg.WriteString(fmt.Sprintf("\r\n--%s--\r\n", boundary))

	adminEmail := strings.Split(getEnv("ADMIN_EMAILS", smtpUser), ",")[0]
	auth := smtp.PlainAuth("", smtpUser, smtpPass, smtpHost)
	if err := smtp.SendMail(smtpHost+":"+smtpPort, auth, smtpUser, []string{strings.TrimSpace(adminEmail)}, msg.Bytes()); err != nil {
		log.Println("Weekly email send error:", err)
	} else {
		log.Println("Weekly volunteer report emailed to", adminEmail)
	}
}

func startWeeklyEmailScheduler() {
	go func() {
		for {
			now := time.Now().UTC()
			daysUntilMonday := (8 - int(now.Weekday())) % 7
			if daysUntilMonday == 0 && (now.Hour() > 3 || (now.Hour() == 3 && now.Minute() >= 30)) {
				daysUntilMonday = 7
			}
			next := time.Date(now.Year(), now.Month(), now.Day()+daysUntilMonday, 3, 30, 0, 0, time.UTC)
			log.Printf("Next weekly report in %s\n", next.Sub(now).Round(time.Minute))
			time.Sleep(next.Sub(now))
			sendExcelEmail()
		}
	}()
}

func sendMilestoneEmail(to, name string, hours int) {
	var subject, body string
	switch hours {
	case 35:
		subject = "Congratulations on 35 Volunteer Hours!"
		body = fmt.Sprintf("Dear %s,\n\nWhat an incredible milestone! You have now clocked 35 volunteer hours with Thuvakkam.\n\nWarm regards,\nTeam Thuvakkam", name)
	case 100:
		subject = "Century Volunteer — 100 Hours of Pure Dedication!"
		body = fmt.Sprintf("Dear %s,\n\nWe are thrilled to celebrate this extraordinary achievement — 100 volunteer hours!\n\nWith deep gratitude,\nTeam Thuvakkam", name)
	default:
		return
	}
	auth := smtp.PlainAuth("", smtpUser, smtpPass, smtpHost)
	msg := []byte("From: " + smtpFrom + "\r\nTo: " + to + "\r\nSubject: " + subject + "\r\nMIME-Version: 1.0\r\nContent-Type: text/plain; charset=UTF-8\r\n\r\n" + body)
	if err := smtp.SendMail(smtpHost+":"+smtpPort, auth, smtpUser, []string{to}, msg); err != nil {
		log.Println("Milestone email error:", err)
	}
}

func checkMilestones(v Volunteer) {
	if v.Hours == 35 || v.Hours == 100 {
		go sendMilestoneEmail(v.Email, v.Name, v.Hours)
	}
}

func parseTemplate() *template.Template {
	return template.Must(
		template.New("dashboard.html").
			Funcs(template.FuncMap{
				"hasSuffix": strings.HasSuffix,
				"sub":       func(a, b int) int { return a - b },
				"pct": func(val, max int) int {
					if max == 0 {
						return 0
					}
					if val >= max {
						return 100
					}
					return val * 100 / max
				},
			}).
			ParseFiles("dashboard.html"),
	)
}

func parseAdminTemplate() *template.Template {
	return template.Must(template.ParseFiles("admin.html"))
}

func render(w http.ResponseWriter, data PageData) {
	if err := parseTemplate().Execute(w, data); err != nil {
		log.Println("Template error:", err)
	}
}

func loginPageHandler(w http.ResponseWriter, r *http.Request) {
	template.Must(template.ParseFiles("new.html")).Execute(w, nil)
}

func loginHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != "POST" {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	email := r.FormValue("login-username")
	password := r.FormValue("login-password")

	var dbPassword string
	err := db.QueryRow("SELECT password FROM volunteers WHERE email=$1", email).Scan(&dbPassword)
	if err != nil || password != dbPassword {
		http.Error(w, "Invalid email or password", http.StatusUnauthorized)
		return
	}

	setSession(w, email)
	http.Redirect(w, r, "/dashboard", http.StatusSeeOther)
}

func registerHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != "POST" {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	name := r.FormValue("reg-username")
	email := r.FormValue("reg-email")
	password := r.FormValue("reg-password")

	_, err := db.Exec(
		"INSERT INTO volunteers (name, email, password, total_hours) VALUES ($1, $2, $3, 0)",
		name, email, password,
	)
	if err != nil {
		http.Error(w, "Could not create account (email may already exist): "+err.Error(), http.StatusInternalServerError)
		return
	}

	setSession(w, email)
	http.Redirect(w, r, "/dashboard", http.StatusSeeOther)
}

func logoutHandler(w http.ResponseWriter, r *http.Request) {
	clearSession(w)
	http.Redirect(w, r, "/", http.StatusSeeOther)
}

func dashboardHandler(w http.ResponseWriter, r *http.Request) {
	v, err := getVolunteer(r)
	if err != nil {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	render(w, PageData{Volunteer: v, Section: "profile"})
}

func profileHandler(w http.ResponseWriter, r *http.Request) {
	v, err := getVolunteer(r)
	if err != nil {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	render(w, PageData{Volunteer: v, Section: "profile"})
}

func achievementsHandler(w http.ResponseWriter, r *http.Request) {
	v, err := getVolunteer(r)
	if err != nil {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	render(w, PageData{Volunteer: v, Section: "achievements"})
}

func eventsHandler(w http.ResponseWriter, r *http.Request) {
	v, err := getVolunteer(r)
	if err != nil {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	render(w, PageData{Volunteer: v, Section: "events"})
}

func feedbackHandler(w http.ResponseWriter, r *http.Request) {
	v, err := getVolunteer(r)
	if err != nil {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	render(w, PageData{Volunteer: v, Section: "feedback"})
}

func uploadPfpHandler(w http.ResponseWriter, r *http.Request) {
	v, err := getVolunteer(r)
	if err != nil {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}
	r.ParseMultipartForm(10 << 20)
	file, handler, err := r.FormFile("profile_pic")
	if err != nil {
		http.Error(w, "Error retrieving file", http.StatusBadRequest)
		return
	}
	defer file.Close()
	filePath := "./uploads/" + handler.Filename
	dst, err := os.Create(filePath)
	if err != nil {
		http.Error(w, "Error saving file", http.StatusInternalServerError)
		return
	}
	defer dst.Close()
	io.Copy(dst, file)
	db.Exec("UPDATE volunteers SET profile_pic=$1 WHERE volunteer_id=$2", filePath, v.ID)
	http.Redirect(w, r, "/profile", http.StatusSeeOther)
}

func adminGuard(w http.ResponseWriter, r *http.Request) (Volunteer, bool) {
	v, err := getVolunteer(r)
	if err != nil {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return Volunteer{}, false
	}
	if !isAdmin(v.Email) {
		http.Error(w, "Access denied", http.StatusForbidden)
		return Volunteer{}, false
	}
	return v, true
}

func adminHandler(w http.ResponseWriter, r *http.Request) {
	_, ok := adminGuard(w, r)
	if !ok {
		return
	}
	volunteers, err := getAllVolunteers()
	if err != nil {
		http.Error(w, "DB error: "+err.Error(), http.StatusInternalServerError)
		return
	}
	parseAdminTemplate().Execute(w, AdminPageData{Volunteers: volunteers})
}

func updateHoursHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != "POST" {
		http.Redirect(w, r, "/admin", http.StatusSeeOther)
		return
	}
	_, ok := adminGuard(w, r)
	if !ok {
		return
	}
	volunteerID, err := strconv.Atoi(r.FormValue("volunteer_id"))
	if err != nil {
		http.Error(w, "Invalid volunteer ID", http.StatusBadRequest)
		return
	}
	hoursToAdd, err := strconv.Atoi(r.FormValue("hours_to_add"))
	if err != nil || hoursToAdd <= 0 {
		http.Error(w, "Invalid hours value", http.StatusBadRequest)
		return
	}
	_, err = db.Exec("UPDATE volunteers SET total_hours = total_hours + $1 WHERE volunteer_id = $2", hoursToAdd, volunteerID)
	if err != nil {
		http.Error(w, "DB error: "+err.Error(), http.StatusInternalServerError)
		return
	}
	var v Volunteer
	db.QueryRow("SELECT volunteer_id, name, email, total_hours FROM volunteers WHERE volunteer_id=$1", volunteerID).
		Scan(&v.ID, &v.Name, &v.Email, &v.Hours)
	checkMilestones(v)
	http.Redirect(w, r, "/admin?flash=updated", http.StatusSeeOther)
}

func exportExcelHandler(w http.ResponseWriter, r *http.Request) {
	_, ok := adminGuard(w, r)
	if !ok {
		return
	}
	volunteers, err := getAllVolunteers()
	if err != nil {
		http.Error(w, "DB error: "+err.Error(), http.StatusInternalServerError)
		return
	}
	f, err := buildExcel(volunteers)
	if err != nil {
		http.Error(w, "Excel error: "+err.Error(), http.StatusInternalServerError)
		return
	}
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", `attachment; filename="volunteers.xlsx"`)
	f.Write(w)
}

func sendReportNowHandler(w http.ResponseWriter, r *http.Request) {
	_, ok := adminGuard(w, r)
	if !ok {
		return
	}
	go sendExcelEmail()
	http.Redirect(w, r, "/admin?flash=emailed", http.StatusSeeOther)
}

func main() {
	godotenv.Load()

	dsn := getEnv("DB_DSN", "")
	var err error
	db, err = sql.Open("postgres", dsn)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()
	if err = db.Ping(); err != nil {
		log.Fatal("Cannot connect to database: ", err)
	}

	os.MkdirAll("uploads", os.ModePerm)
	startWeeklyEmailScheduler()

	http.HandleFunc("/imag.jpg", func(w http.ResponseWriter, r *http.Request) { http.ServeFile(w, r, "imag.jpg") })
	http.HandleFunc("/logo.png", func(w http.ResponseWriter, r *http.Request) { http.ServeFile(w, r, "logo.png") })
	http.Handle("/uploads/", http.StripPrefix("/uploads/", http.FileServer(http.Dir("uploads"))))

	http.HandleFunc("/", loginPageHandler)
	http.HandleFunc("/login", loginHandler)
	http.HandleFunc("/register", registerHandler)
	http.HandleFunc("/logout", logoutHandler)
	http.HandleFunc("/dashboard", dashboardHandler)
	http.HandleFunc("/profile", profileHandler)
	http.HandleFunc("/achievements", achievementsHandler)
	http.HandleFunc("/events", eventsHandler)
	http.HandleFunc("/feedback", feedbackHandler)
	http.HandleFunc("/upload-pfp", uploadPfpHandler)

	http.HandleFunc("/admin", adminHandler)
	http.HandleFunc("/admin/update-hours", updateHoursHandler)
	http.HandleFunc("/admin/send-report", sendReportNowHandler)
	http.HandleFunc("/export-excel", exportExcelHandler)

	port := getEnv("PORT", "8080")
	fmt.Println("Server running at http://localhost:" + port)
	log.Fatal(http.ListenAndServe(":"+port, nil))
}
