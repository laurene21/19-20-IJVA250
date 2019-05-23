package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Set;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();

        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        writer.println("Id" + ";" + "Nom" + ";" + "Prenom" + ";" + "Date de Naissance" + ";" + "Age");

        for( Client client : allClients ){
            writer.println( client.getId() + ";" + client.getNom() + ";\"" + client.getPrenom() + "\";" +
                    client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")) + ";" + (now.getYear() - client.getDateNaissance().getYear()));
        }

    }

    @GetMapping("/clients/xlsx")
    public void clientXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        Cell cellNom = headerRow.createCell(1);
        Cell cellPrenom = headerRow.createCell(2);
        Cell cellDn = headerRow.createCell(3);
        Cell cellAge = headerRow.createCell(4);
        cellId.setCellValue("Id");
        cellNom.setCellValue("Nom");
        cellPrenom.setCellValue("Prenom");
        cellDn.setCellValue("Date de naissance");
        cellAge.setCellValue("Age");


        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        int i = 1;
        for( Client client : allClients ){
            Row newRow = sheet.createRow(i);
            newRow.createCell(0).setCellValue(client.getId());
            newRow.createCell(1).setCellValue(client.getNom());
            newRow.createCell(2).setCellValue(client.getPrenom());
            newRow.createCell(3).setCellValue(client.getDateNaissance().toString());
            newRow.createCell(4).setCellValue((now.getYear() - client.getDateNaissance().getYear()));
            i++;
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{id}/factures/xlsx")
    public void factureXLSXByClient(@PathVariable("id") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures-client-" + clientId + ".xlsx\"");
        List<Facture> factures = factureService.findFactureClient(clientId);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Facture");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellTotal = headerRow.createCell(1);
        cellTotal.setCellValue("Prix Total");

        int iRow = 1;
        for (Facture facture : factures) {
            Row row = sheet.createRow(iRow);

            Cell id = row.createCell(0);
            id.setCellValue(facture.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(facture.getTotal());

            iRow = iRow + 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void getAllFactures(HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");

        Workbook workbook = new XSSFWorkbook();

        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        for(Client client : allClients){

            Sheet sheetClient = workbook.createSheet(client.getNom());
            sheetClient.createRow(0).createCell(0).setCellValue(client.getNom());
            sheetClient.createRow(1).createCell(0).setCellValue(client.getPrenom());
            sheetClient.createRow(2).createCell(0).setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));


            List<Facture> factures = factureService.findFactureClient(client.getId());

            //Page facture
            for(Facture facture : factures){

                Sheet sheetFacture = workbook.createSheet("Facture" + facture.getId());
                Row headerRow = sheetFacture.createRow(0);
                headerRow.createCell(0).setCellValue("Nom de l'article");
                headerRow.createCell(1).setCellValue("Quantité");
                headerRow.createCell(2).setCellValue("Prix unitaire");
                headerRow.createCell(3).setCellValue("Sous-total");

                //Ligne par facture
                int i = 1;
                for(LigneFacture ligneFacture : facture.getLigneFactures()){
                    Row newRow = sheetFacture.createRow(i);
                    newRow.createCell(0).setCellValue(ligneFacture.getArticle().getLibelle());
                    newRow.createCell(1).setCellValue(ligneFacture.getQuantite());
                    newRow.createCell(2).setCellValue(ligneFacture.getArticle().getPrix());
                    newRow.createCell(3).setCellValue(ligneFacture.getSousTotal());
                    i++;
                }

                //Total
                Row totalRow = sheetFacture.createRow(i);
                Cell totalCell = totalRow.createCell(2);
                totalCell.setCellValue("TOTAL");
                Cell totalCell2 = totalRow.createCell(3);
                totalCell2.setCellValue(facture.getTotal());

                //Style
                CellStyle cellStyle  = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                font.setColor(IndexedColors.RED.getIndex());

                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);

                cellStyle.setFont(font);

                totalCell.setCellStyle(cellStyle);
                totalCell2.setCellStyle(cellStyle);
            }

        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }

}
