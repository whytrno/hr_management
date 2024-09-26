# Copyright (c) 2024, Wahyu Triono and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from frappe.utils.file_manager import save_file

class SuratTugas(Document):
	def before_save(self):
		karyawan_items = self.get_karyawan_data();

		nama_surat = f'{self.no_surat.replace('/', '-')}.docx'
		# file = frappe.utils.get_site_path('private', 'files', 'test.txt')

		generated_docx = self.create_document(nama_surat, karyawan_items)
		self.upload_document_to_file_manager(nama_surat, generated_docx)

	def get_karyawan_data(self):
		karyawan_items = []

		# Ambil data dari child table `karyawan` yang ada di DocType SuratTugas
		for item in self.karyawan:
			karyawan_doc = frappe.get_doc('User', item.user_email)  # asumsikan ada field karyawan_id
			karyawan_items.append({
				'no': item.idx,
				'nama': karyawan_doc.full_name,
				'nrp': karyawan_doc.nrp,
				'jabatan': karyawan_doc.jabatan,
				'tk': self.tanggal_keberangkatan,
				'foto_ktp': karyawan_doc.foto_ktp,
				'foto_vaksin': karyawan_doc.foto_vaksin,
			})

		return karyawan_items

	def create_document(self, nama_surat, karyawan_items):
		template_path = frappe.utils.get_site_path('private', 'templates', 'st_kelompok.docx')
		doc = DocxTemplate(template_path)
		tanggal = frappe.utils.formatdate(frappe.utils.nowdate(), "dddd, dd MMMM yyyy")

		table_context = []
		for karyawan in karyawan_items:
			foto_ktp_path = None
			foto_vaksin_path = None

			if karyawan['foto_ktp']:
				foto_ktp_name = karyawan['foto_ktp'].split('/')[-1]
				foto_ktp_path = frappe.utils.get_site_path('private', 'files', foto_ktp_name)

			if karyawan['foto_vaksin']:
				foto_vaksin_name = karyawan['foto_vaksin'].split('/')[-1]
				foto_vaksin_path = frappe.utils.get_site_path('private', 'files', foto_vaksin_name)

			table_context.append({
				'no': karyawan['no'],
				'nama': karyawan['nama'],
				'nrp': karyawan['nrp'],
				'jabatan': karyawan['jabatan'],
				'tk': karyawan['tk'],
				'foto_ktp': InlineImage(doc, foto_ktp_path, width=Mm(20), height=Mm(20)) if foto_ktp_path else None,
				'foto_vaksin': InlineImage(doc, foto_vaksin_path, width=Mm(20), height=Mm(20)) if foto_vaksin_path else None,
			})

		# Siapkan context untuk template
		context = {
			'no_surat': self.no_surat,
			'keperluan': self.keperluan,
			'lokasi_site': self.lokasi_site,
			'tanggal': tanggal,
			'karyawan': table_context
		}

		# Render template
		doc.render(context)

		# Simpan dokumen yang dihasilkan
		docx_file_path = frappe.utils.get_site_path('private', 'files', nama_surat)
		doc.save(docx_file_path)

		return docx_file_path

		# pdf_file_path = frappe.utils.get_site_path('private', 'files', 'surat_keterangan.pdf')
		# pypandoc.convert_file(docx_file_path, 'pdf', outputfile=pdf_file_path)

	def create_folders_if_not_exist(self, base_path, folder_path):
		# Pisahkan path menjadi komponen folder
		folders = folder_path.split('/')

		# Variable untuk melacak full path dari folder
		current_folder = base_path

		# Loop melalui setiap folder di path
		for folder in folders:
			current_folder = f"{current_folder}/{folder}"

			# Jika folder belum ada, buat folder baru
			if not frappe.db.exists('File', current_folder):
				folder_doc = frappe.get_doc({
					"doctype": "File",
					"file_name": folder,
					"folder": base_path if current_folder == f"{base_path}/{folder}" else current_folder.rsplit('/', 1)[0],
					"is_folder": 1,
					"is_private": 1
				})
				folder_doc.save()
				frappe.logger().info(f"Folder '{folder}' created in '{current_folder}'")

	def upload_document_to_file_manager(self, nama_surat, file):
		with open(file, 'rb') as f:
			file_data = f.read()

		# Get the current year and month in Roman numeral format
		year_now = frappe.utils.nowdate().split('-')[0]
		month_now_roman = frappe.utils.formatdate(frappe.utils.nowdate(), "MM")

		# Define the full folder path based on your structure
		folder_structure = f"Surat/Surat Tugas/{year_now}/{month_now_roman}"

		# Create folders if not exist
		self.create_folders_if_not_exist("Home", folder_structure)

		# Final folder path after folder creation
		final_folder = f'Home/{folder_structure}'

		# Create the new file in the final folder
		_file = frappe.get_doc({
			"doctype": "File",
			"file_name": nama_surat,
			"folder": final_folder,
			"content": file_data,
			"is_private": 1,
			"ignore_duplicate_entry_error": True
		})
		_file.save()

		# Get the file URL (this is the public URL if the file is public, otherwise it's a private path)
		file_url = _file.file_url

		# Save the file URL or path back to the document
		self.file_url = file_url  # Assumes 'file_url' field exists on the Surat Tugas doctype

		# Optionally log or debug
		frappe.logger().info(f"File {nama_surat} saved successfully with URL {file_url}")
