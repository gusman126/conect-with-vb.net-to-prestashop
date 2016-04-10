Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Net

Public Class prestashop
    Dim conn As New MySqlConnection '' conexion al servidor y configuracion


    '' -------------------- DATOS DEL SERVIDOR MYSQL - RECORDAD QUE SE DEBE TENER LA OPCION DE MYSQL REMOTO ACTIVADO EN EL SERVIDOR

    Dim DatabaseName As String
    Dim server As String
    Dim userName As String
    Dim password As String
    Dim fecha As String


    Dim id_idioma As Integer



    '' datos FTP y variables para subir las imagenes

    Dim server_ftp, user_ftp, passw_ftp As String ' datos de FTP para subir las fotos
    Dim idimage, id_image_extra(8) As Integer ' id de la imagen principal y las 8 imagenes extras
    Public imagen_subir, imagenprincipal, imagen_subir_extra(8), imagen_extra(8) As String ' string de las imagenes del explorador de archivos, la ruta y el nombre 

    Dim path_ftp_image As String '  el camino o path que genera con el ID de la imagen para subirla al servidor

    '' Datos del producto, nombre, id, numero de categorias que van , la categoria donde ira ademas de la de home etc..

    '' variables de producto
    Dim idproducto As Integer

    Dim precio_producto, nombre_producto, referencia_producto, codigoEAN, descripcion, descripcion_larga As String



    ''vriablea de categoria
    Dim categorias(9999) As Integer
    Dim categoria As String
    Dim id_categoria As Integer

    ''variables de stock

    Dim stock_total As Integer


    ''variables y valores de opciones en este caso color y talla y los ID de estos valores, el precio extra de esa combinacion

    Dim color_val, talla_val As String
    Dim stock_attributo As Integer
    Dim id_color, id_talla As Integer
    Dim precio_extra As Integer

    '' variables de attributo al añadir y sus combinaciones

    Dim ultimo_attributo As Integer

    '' datos de categirias
    Public name_category As String


    ' DATOS DE ATTRIBUTOS
    Dim attr_principal(1000), attr_secundario(1000) As String



    ' ABRIENDO EL FORMULARIO PRIMERAS ORDENES

    Private Sub prestashop_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' ----- DATOS DEL SERVIDOR MYSQL
        'DatabaseName = "Data base name"
        'server = "domain.com"
        'userName = "user_db"
        'password = "pass_db"

        DatabaseName = "Data base name"
        server = "domain.com"
        userName = "user_db"
        password = "pass_db"


        connect()


        fecha = "CURRENT_TIMESTAMP"




        

        ' ------------- DATOS DEL SERVIDOR ftp debemos añadir una cuenta que vaya al raiz de la tienda, si nuestra tienda es dominio.com/tienda .. 
        ' ------------- debemos añadir un usuario ftp que acceda directamente a /tienda

        'server_ftp = "domain.com"
        'user_ftp = "user@domain.com"
        'passw_ftp = "uug7onr(PL2g08^W"


        server_ftp = "domain.com"
        user_ftp = "user@domain.com"
        passw_ftp = "pass_ftp"


       
        ' ----- BUSCAMOS EL ID DE ESPAÑOL

        id_lang("español")

        ' ---- AÑADIMOS LOS ATRIBUTOS DE LA BD AL PROGRAMA
        read_attributes_group()

        '' AÑADIMOS LAS CATEGORIAS DE LA BD AL PROGRAMA
        add_categories()
    End Sub

    ' --------------------------- CONEXION CON BASE DE DATOS

    Public Sub connect() ' conecta con la base de datos
        If Not conn Is Nothing Then conn.Close()
        conn.ConnectionString = String.Format("server={0}; user id={1}; password={2}; database={3}; pooling=false", server, userName, password, DatabaseName)
        Try
            conn.Open()

            MsgBox("Connexion con el servidor OK")

        Catch ex As Exception
            MsgBox("Error al conectar con el servidor!, compruebe que tiene acceso a internet" & ex.Message)
        End Try
        conn.Close()
    End Sub

    ' BUSCAMOS EL ID DEL IDIMA
    Private Sub id_lang(languaje As String)
        Dim cm_id_lang As New MySqlCommand(String.Format("SELECT `id_lang` FROM `ps_lang` WHERE `name` LIKE '%" & languaje & "%'"), conn)

        Try
            conn.Open()
            id_idioma = cm_id_lang.ExecuteScalar
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    ' ------------------------------------- ATRIBUTOS

    ' - Añadiendo atributos a los combobox
    Private Sub read_attributes_group()
        Dim name_attributes_group As New MySqlCommand(String.Format("SELECT * FROM `ps_attribute_group` join `ps_attribute_group_lang` on `ps_attribute_group`.`id_attribute_group` = `ps_attribute_group_lang`.`id_attribute_group` WHERE `id_lang` = " & id_idioma & "  "), conn)
        Dim re_attr_g As MySqlDataReader

        conn.Open()
        re_attr_g = name_attributes_group.ExecuteReader
        Do While re_attr_g.Read
            c_attr_group_p.Items.Add(re_attr_g("name"))
            c_attr_group_s.Items.Add(re_attr_g("name"))
        Loop
        conn.Close()
    End Sub
    Private Sub c_attr_group_p_SelectedIndexChanged(sender As Object, e As EventArgs) Handles c_attr_group_p.SelectedIndexChanged
        Dim cm_id_attr_group As New MySqlCommand(String.Format("SELECT `id_attribute_group` FROM `ps_attribute_group_lang` WHERE `name` = '" & c_attr_group_p.SelectedItem & "'"), conn)
        Dim id_attr_group As Integer
        conn.Open()

        id_attr_group = cm_id_attr_group.ExecuteScalar
        conn.Close()

        read_attributes_p(id_attr_group)

    End Sub

    Private Sub read_attributes_p(parent)
        Dim name_attributes As New MySqlCommand(String.Format("SELECT * FROM `ps_attribute` join `ps_attribute_lang` on `ps_attribute`.`id_attribute` = `ps_attribute_lang`.`id_attribute` WHERE `ps_attribute`.`id_attribute_group` = " & parent & " and `ps_attribute_lang`.`id_lang` = " & id_idioma & ""), conn)
        Dim re_attributes As MySqlDataReader
        tree_attributes_p.Nodes.Clear()

        conn.Open()
        re_attributes = name_attributes.ExecuteReader

        Do While re_attributes.Read
            tree_attributes_p.Nodes.Add(re_attributes("id_attribute"), re_attributes("name"))
        Loop
        conn.Close()

    End Sub

    Private Sub c_attr_group_s_SelectedIndexChanged(sender As Object, e As EventArgs) Handles c_attr_group_s.SelectedIndexChanged
        Dim cm_id_attr_group As New MySqlCommand(String.Format("SELECT `id_attribute_group` FROM `ps_attribute_group_lang` WHERE `name` = '" & c_attr_group_s.SelectedItem & "'"), conn)
        Dim id_attr_group As Integer
        conn.Open()

        id_attr_group = cm_id_attr_group.ExecuteScalar
        conn.Close()

        read_attributes_s(id_attr_group)
    End Sub
    Private Sub read_attributes_s(parent)
        Dim name_attributes As New MySqlCommand(String.Format("SELECT * FROM `ps_attribute` join `ps_attribute_lang` on `ps_attribute`.`id_attribute` = `ps_attribute_lang`.`id_attribute` WHERE `ps_attribute`.`id_attribute_group` = " & parent & " and `ps_attribute_lang`.`id_lang` = " & id_idioma & ""), conn)
        Dim re_attributes As MySqlDataReader
        tree_attributes_s.Nodes.Clear()

        conn.Open()
        re_attributes = name_attributes.ExecuteReader

        Do While re_attributes.Read
            tree_attributes_s.Nodes.Add(re_attributes("id_attribute"), re_attributes("name"))
        Loop
        conn.Close()

    End Sub

    ' --------------------------AÑADIENDO LAS VARIABLES DE ATTRIBUTOS PRINCIPAL, SECUNDARIOS, STOCK Y AÑADIENDO A LA REGILLA


    Private Sub add_attr_Click(sender As Object, e As EventArgs) Handles add_attr.Click
        'opciones.Rows.Clear()
        If opciones.Columns.Count = 0 Then
            opciones.Columns.Add(c_attr_group_p.Text.ToLower, c_attr_group_p.Text)
            opciones.Columns.Add(c_attr_group_s.Text.ToLower, c_attr_group_s.Text)
            opciones.Columns.Add("stock", "Stock")

        Else

        End If
        

        Dim list_prin, list_sec As New ArrayList


        For Each chk_prin As TreeNode In tree_attributes_p.Nodes
            If chk_prin.Checked = True Then
                list_prin.Add(chk_prin.Text)
            End If
        Next

        For Each chk_sec As TreeNode In tree_attributes_s.Nodes
            If chk_sec.Checked = True Then
                list_sec.Add(chk_sec.Text)
            End If
        Next
       
        Dim comp, coms As String

        For Each comp In list_prin

            For Each coms In list_sec
                opciones.Rows.Add(comp, coms, stock_attr.Value)

            Next

        Next


    End Sub


    ' ------------------------------CATEGORIAS

   
    Private Sub add_categories()

        ' busca y añade al arbol las categorias


        Dim cat_principales As New MySqlCommand(String.Format("SELECT * FROM `ps_category` join `ps_category_lang` on  `ps_category`.`id_category` = `ps_category_lang`.`id_category` where `ps_category_lang`.`id_lang` = " & id_idioma & ""), conn)
        Dim lis_principales As MySqlDataReader
        conn.Open()
        lis_principales = cat_principales.ExecuteReader

        Do While lis_principales.Read
            Try
                If lis_principales("id_parent") < 2 Then ' si es la categoria 1 o 2 son las principales las añade normalmente
                    listado_categorias.Nodes.Add(lis_principales("id_category"), lis_principales("name"))
                Else ' las demas las añade en el arbol con su correspondiente categoria padre
                    listado_categorias.Nodes.Find(lis_principales("id_parent"), True).First.Nodes.Add(lis_principales("id_category"), lis_principales("name"))
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Loop

        conn.Close()

        listado_categorias.ExpandAll()


    End Sub

    ' ------------- CADA VEZ QUE ELEGIMOS UNA CATEGORIA LA AÑADE A LA VARIABLE
    Private Sub listado_categorias_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles listado_categorias.AfterCheck
        If e.Node.Checked = True Then
            categorias(e.Node.Name) = e.Node.Name
        Else
            categorias(e.Node.Name) = ""
        End If
    End Sub

    ' -------------------- AÑADIR EL PRODUCTO Y TODO EL CODIGO SIGUIENTE

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If tnameweb.Text = Nothing Or tdescripweb.Text = Nothing Or tpvpweb.Text = Nothing Or timagen_principal.Text = Nothing Then
            MsgBox("Debe rellenar los textos de nombre, descripción y precio. Ademas de añadir una foto principal",MsgBoxStyle.Critical)
        Else
            Try
                conn.Open()
                insert()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            conn.Close()
        End If

        

    End Sub


    ' ------------------------------------- INSERTAR PRODUCTO

    Private Sub insert()

        nombre_producto = tnameweb.Text
        referencia_producto = tnameweb.Text
        descripcion = tdescripweb.Text
        descripcion_larga = ""

        codigoEAN = ""

        precio_producto = tpvpweb.Text



        ' añade el producto  
        Dim add_prod As New MySqlCommand(String.Format("insert into ps_product (id_category_default,id_shop_default, ean13, quantity,minimal_quantity,price,reference,available_date, date_add, advanced_stock_management,pack_stock_type)" & _
                                                       " Values (2,1,'" & codigoEAN & "',0,1,'" & precio_producto.ToString & "','" & referencia_producto & "', " & fecha & "," & fecha & ",1," & id_idioma & ")"), conn)
        add_prod.ExecuteNonQuery()

        ' busca el producto añadido
        Dim cmd_idproducto As New MySqlCommand(String.Format("SELECT * FROM ps_product WHERE id_product = (SELECT MAX(id_product) FROM ps_product)"), conn)


        idproducto = cmd_idproducto.ExecuteScalar
        MsgBox("ID del producto es : " & idproducto)

        ' Añade la descripcion

        Dim add_descrip As New MySqlCommand(String.Format("INSERT INTO `ps_product_lang`(`id_product`, `id_shop`, `id_lang`, `description`, `description_short`,`link_rewrite`, `name`) " & _
                                                          " Values ( " & idproducto & ",1," & id_idioma & ", ' " & descripcion_larga & " ','" & descripcion & "','" & StrConv(nombre_producto.Replace(" ", "-"), VbStrConv.Lowercase) & "','" & nombre_producto & "') "), conn)



        add_descrip.ExecuteNonQuery()



        ' añade el producto a la tienda

        Dim add_shop As New MySqlCommand(String.Format("INSERT INTO `ps_product_shop`(`id_product`, `id_shop`, `id_category_default`, `id_tax_rules_group`, `on_sale`, `online_only`, `ecotax`, `minimal_quantity`, `price`, `active`,`available_for_order`, `available_date`, `condition`, `show_price`, `indexed`, `visibility`, `advanced_stock_management`, `date_add`,`pack_stock_type`) " & _
                                                        "VALUES (" & idproducto & ",1,2,53,0,0,0,1,'" & precio_producto.ToString & "',1,1, " & fecha & ",'new',1,1,'both',1," & fecha & "," & id_idioma & ") "), conn)
        add_shop.ExecuteNonQuery()


        

        ' Compruba cuantas categorias tiene asignado el producto  y lo añade

        For cat As Integer = 0 To categorias.Length - 1
            If categorias(cat) = 0 Then
            Else
                Dim add_cat_product As New MySqlCommand(String.Format("INSERT INTO `ps_category_product`(`id_category`, `id_product`, `position`) " & _
                                                         "VALUES (" & categorias(cat) & "," & idproducto & "," & id_idioma & ") "), conn)
                add_cat_product.ExecuteNonQuery()
            End If
            
        Next




        '' antes de añadir el stock del attributo debemos poner en 0 el total de ese atributo y sumar el stock de cada par de attributos

        stock_total = 0

        '' debemos poner una combinacion principal
        Dim combinacion_principal As Integer



        ''''''''''' Esta orden debe estar en un FOR X , para que lo haga X veces para todas las combinaciones de Color + Talla


        For x As Integer = 0 To opciones.RowCount - 2 '' he puesto -2 porque se quedaba una fila al final en blanco
            opciones.Rows(x).Selected = True
            opciones.Refresh()

            combinacion_principal = x



            ' Añadimos el stock de cada producto por pares color + talla, por ejemplo si de fucsia y talla hay 10, pondremos 10

            color_val = opciones.Item(opciones.Columns(c_attr_group_p.Text.ToLower).Index, x).Value
            talla_val = opciones.Item(opciones.Columns(c_attr_group_s.Text.ToLower).Index, x).Value
            stock_attributo = opciones.Item(opciones.Columns("stock").Index, x).Value

            ' Pero primero buscamos los Id de color y talla

            Dim search_id_color As New MySqlCommand(String.Format("select id_attribute from ps_attribute_lang where name = '" & color_val & "'"), conn)
            Dim search_id_talla As New MySqlCommand(String.Format("select id_attribute from ps_attribute_lang where name = '" & talla_val & "'"), conn)

            id_color = search_id_color.ExecuteScalar
            id_talla = search_id_talla.ExecuteScalar

            If id_color = 0 Or id_talla = 0 Then
                MsgBox("No existe esta opcion")
            End If

            '' buscamos el ultimo attributo
            Dim last_attribute As New MySqlCommand(String.Format("select id_product_attribute from ps_product_attribute Order by id_product_attribute DESC limit 1"), conn)

            ultimo_attributo = last_attribute.ExecuteScalar + 1



            '' creamos un EAN y una Referencia unica para esta combinacion
            Dim referencia, ean As String
            referencia = referencia_producto & Rnd(100).ToString & color_val & talla_val


            '' aqui deberia buscar en la base de datos el EAN generado por el programa o generarlo
            ean = "0"


            '' añadimos el attrbuto si tiene un precio extra lo añadiremos a 'price'


            precio_extra = 0.0


            Dim add_attribute As New MySqlCommand(String.Format("INSERT INTO `ps_product_attribute` (`id_product_attribute`,`id_product`, `reference`, `ean13`, `price`, `quantity`,`minimal_quantity`, `available_date`,`default_on`) " & _
                                                               "values ( " & ultimo_attributo & ", " & idproducto & ",'" & referencia_producto & "','" & ean & "', " & precio_extra & "," & stock_attributo & ",1, " & fecha & "," & combinacion_principal & " )"), conn)

            add_attribute.ExecuteNonQuery()


            '' Añadimos esta combinacion al almacen en este caso solo hay un almacen con ID 2
            Dim add_almacen As New MySqlCommand(String.Format(" INSERT INTO `ps_warehouse_product_location` (`id_product`, `id_product_attribute`, `id_warehouse`, `location`) " & _
                                                              " Values ( " & idproducto & ", " & ultimo_attributo & ",2,'')"), conn)
            add_almacen.ExecuteNonQuery()


            ''''' añadimos este attributo a la tienda
            Dim add_atribute_shop As New MySqlCommand(String.Format("INSERT INTO `ps_product_attribute_shop`(`id_product`, `id_product_attribute`, `id_shop`, `price`, `minimal_quantity`, `available_date`,`default_on`) " & _
                                                                    "values (" & idproducto & "," & ultimo_attributo & ",1, " & precio_extra & ",1, " & fecha & "," & combinacion_principal & ")"), conn)

            add_atribute_shop.ExecuteNonQuery()



            ''Se añaden las combinaciones, debemos recordar que son dos opciones color + talla por cada atributo

            '' Ahora añadimos a ese atributo el color fucsia

            Dim add_color As New MySqlCommand(String.Format("insert into ps_product_attribute_combination (id_attribute, id_product_attribute) values (" & id_color & ", " & ultimo_attributo & ")"), conn)
            add_color.ExecuteNonQuery()

            '' ahora añadimos al mismo attributo la talla

            Dim add_talla As New MySqlCommand(String.Format("insert into ps_product_attribute_combination (id_attribute, id_product_attribute) values (" & id_talla & ", " & ultimo_attributo & ")"), conn)
            add_talla.ExecuteNonQuery()




            '' Ahora debemos añadir el stock de cada attributo y combinacion, el disponible y el stock en este momento, pero ademas añadimos este stock al total de ese attributo

            '' Stock disponible
            Dim add_stock_available As New MySqlCommand(String.Format("INSERT INTO `ps_stock_available` (`id_product`, `id_product_attribute`, `id_shop`, `quantity`, `out_of_stock`) " & _
                                                        "Values (" & idproducto & ", " & ultimo_attributo + 1 & ",1, " & stock_attributo & ",2)"), conn)
            add_stock_available.ExecuteNonQuery()



            '' Stock en este momento

            Dim add_stock As New MySqlCommand(String.Format("INSERT INTO `ps_stock` (`id_warehouse`, `id_product`, `id_product_attribute`, `reference`, `ean13`, `physical_quantity`, `usable_quantity`, `price_te`) " & _
                                                            "values(2," & idproducto & ", " & ultimo_attributo & ",'" & referencia_producto & "', '" & ean & "'," & stock_attributo & ", " & stock_attributo & "," & precio_producto & ") "), conn)

            add_stock.ExecuteNonQuery()




            '' Añadirmos un movimiento de stock al almacen 1, buscnando antes cual es el ultimo 

            Dim last_stock As New MySqlCommand(String.Format("select id_stock from ps_stock Order by id_stock DESC limit 1"), conn)

            Dim ultimo_stock As Integer

            ultimo_stock = last_stock.ExecuteScalar

            Dim add_mov_stock As New MySqlCommand(String.Format("INSERT INTO `ps_stock_mvt`( `id_stock`, `id_stock_mvt_reason`, `employee_lastname`, `physical_quantity`, `date_add`, `sign`, `price_te`) " & _
                                                                " VALUES (" & ultimo_stock & ",1,'TPV'," & stock_attributo & "," & fecha & ",1," & precio_producto & ")"), conn)

            add_mov_stock.ExecuteNonQuery()

            '' sumamos el stock de esa combinacion al stock total
            stock_total = stock_total + stock_attributo

            ' el primero es el principal, los siguientes no



        Next
        ' FINAL DEL FOR y ya se han añadido todas las combinaciones




        ' Sumamos todo el stock al final de ese producto sin el numero de attributo, poniendo 0
        Dim add_stock_total As New MySqlCommand(String.Format("INSERT INTO `ps_stock_available` (`id_product`, `id_product_attribute`, `id_shop`, `quantity`,`depends_on_stock`, `out_of_stock`) " & _
                                                              " values (" & idproducto & ",0,1," & stock_total & ",1,2) "), conn)

        add_stock_total.ExecuteNonQuery()

        ' actualizar el total del stock del producto

        Dim upd_stock_producto As New MySqlCommand(String.Format("UPDATE `ps_product` SET `quantity`=" & stock_total & "  WHERE `id_product` = " & idproducto & ""), conn)
        upd_stock_producto.ExecuteNonQuery()


        ' cerramos la base de datos por ahora
        conn.Close()

        Try
            '' subimos la imagen principal 
            BackgroundWorker1.RunWorkerAsync()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    '  ------------------------ SEGUIMOS AÑADIENDO EL PRODUCTO AHORA IMAGENES Y EN "BACKGROUND"

    Private Sub add_images(id_producto, position, principal)
        '''' IMAGEN PRINCIPAL

        ' comprobamos la ultima id_image y sumamos 1 para añadirla a la base de datos

        Dim cmd_idimage As New MySqlCommand(String.Format("SELECT * FROM ps_image WHERE id_image = (SELECT MAX(id_image) FROM ps_image)"), conn)
        idimage = cmd_idimage.ExecuteScalar + 1

        MsgBox("La id de la imagen es: " & idimage)


        'añadimos la id imagen al producto
        Dim add_image As New MySqlCommand(String.Format("insert into  `ps_image`(`id_image`, `id_product`, `position`, `cover`) VALUES (" & idimage & "," & id_producto & "," & position & ", " & principal & ")"), conn)
        add_image.ExecuteNonQuery()


        'asignamos la imagen a los idiomas el id 3 = español
        Dim image_to_lang As New MySqlCommand(String.Format("INSERT INTO `ps_image_lang`(`id_image`, `id_lang`,`legend`) VALUES (" & idimage & "," & id_idioma & ",'" & StrConv(nombre_producto.Replace(" ", "-"), VbStrConv.Lowercase) & "')"), conn)
        image_to_lang.ExecuteNonQuery()

        ' asignamos esa id_image a la tienda y añadimos el id_producto LA ID_SHOP ES 1, cover = es principal

        Dim image_to_product As New MySqlCommand(String.Format("INSERT INTO `ps_image_shop`(`id_product`, `id_image`, `id_shop`, `cover`) VALUES (" & id_producto & "," & idimage & ",1," & principal & ")"), conn)
        image_to_product.ExecuteNonQuery()

    End Sub


    Public Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ' aqui ponemos la orden que subira la imagen principal al servidor
        conn.Open()
        '' añadimos a la base de datos una imagen y nos devuelve el ID

        add_images(idproducto, 1, 1)

        make_ftp_folder(idimage.ToString)
        subir_ftp_imagen(path_ftp_image, timagen_principal.Text)


        '' ahora las imagenes extras
        If timagen_extra1.Text = "" Then
        Else
            add_images(idproducto, 2, 0)
            make_ftp_folder(idimage.ToString)
            subir_ftp_imagen(path_ftp_image, timagen_extra1.Text)

        End If



    End Sub


    Public Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        progreso_img_pricipal.Value = e.ProgressPercentage
        lbl_img_principal.Text = e.ProgressPercentage.ToString & " %"

    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        MsgBox("fin")
        conn.Close()

    End Sub

    ' ----------------- CREANDO LAS CARPETAS EN EL SERVIDOR

    Private Sub make_ftp_folder(id_image As String)
        '        MsgBox("FTP")

        ' todo este codigo es copia de uno que tenia, por lo cual funciona pero puede que no sea necesario todo el codigo, mejor dejarlo
        path_ftp_image = ""
        For y As Integer = 0 To id_image.Length - 1

            Dim x As String
            x = ""
            x = id_image.Substring(y, 1)
            path_ftp_image = path_ftp_image & "/" & x

            '' Comprobar que exista la carpeta de categorias.
            Dim add_folder As Net.FtpWebRequest = CType(FtpWebRequest.Create("FTP://" & server_ftp & "/img/p" & path_ftp_image), FtpWebRequest)
            add_folder.Credentials = New System.Net.NetworkCredential(user_ftp, passw_ftp)

            '' crea la carpeta
            add_folder.Method = WebRequestMethods.Ftp.MakeDirectory
            Try
                Using response As FtpWebResponse = DirectCast(add_folder.GetResponse(), FtpWebResponse)

                End Using
            Catch ex As Exception
            End Try
        Next
    End Sub


    ' ---------------------- SUBIENDO LAS IMAGENES

    Private Sub subir_ftp_imagen(path, imagen_up)

        '' Sube la imagen principal a la carpeta de su ID
        Try
            Dim conect_ftp As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("FTP://" & server_ftp & "/img/p" & path & "/" & idimage & ".jpg"), System.Net.FtpWebRequest)
            conect_ftp.Method = System.Net.WebRequestMethods.Ftp.UploadFile
            conect_ftp.Credentials = New System.Net.NetworkCredential(user_ftp, passw_ftp)
            conect_ftp.UseBinary = True

            Dim fichero As Byte() = System.IO.File.ReadAllBytes(imagen_up)
            Dim requeststream As System.IO.Stream = conect_ftp.GetRequestStream

            For offset As Integer = 0 To fichero.Length Step 1024
                BackgroundWorker1.ReportProgress(CType(offset * progreso_img_pricipal.Maximum / fichero.Length, Integer))

                Dim chsize As Integer = fichero.Length - offset
                If chsize > 1024 Then chsize = 1024
                requeststream.Write(fichero, offset, chsize)


            Next
            requeststream.Close()
            requeststream.Dispose()


        Catch ex As Exception
            MsgBox("error al subir la imagen : " & "/img/p" & path & "/" & idimage & ".jpg" & "/br" & ex.Message)

        End Try
    End Sub

    



    '' explorador de ficheros y pone la imagen en el recuadro
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Dim imagen As New OpenFileDialog
        If imagen.ShowDialog = Windows.Forms.DialogResult.OK Then
            timagen_principal.Text = imagen.FileName
            imagenprincipal = imagen.FileName
            imagen_subir = IO.Path.GetFileName(imagen.FileName)
        End If
        Dim bm As New Bitmap(imagenprincipal)

        imagen_principal.Image = bm
        imagen_principal.SizeMode = PictureBoxSizeMode.StretchImage

    End Sub

    ' ------------------------ BOTONES PARA AÑADIR IMAGENES


    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim imagen As New OpenFileDialog
        If imagen.ShowDialog = Windows.Forms.DialogResult.OK Then
            timagen_extra1.Text = imagen.FileName
            imagen_extra(1) = imagen.FileName
            imagen_subir_extra(1) = IO.Path.GetFileName(imagen.FileName)
        End If
        Dim bm As New Bitmap(imagen_extra(1))
        box_img_1.Image = bm
        box_img_1.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

   

    




    'Private Sub subir_ftp_imagen_home(path, imagen_up)

    '    '' Sube la imagen principal a la carpeta de su ID
    '    Try
    '        Dim conect_ftp As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("FTP://" & server_ftp & "/img/p" & path & "/" & idimage & "-home_default.jpg"), System.Net.FtpWebRequest)
    '        conect_ftp.Method = System.Net.WebRequestMethods.Ftp.UploadFile
    '        conect_ftp.Credentials = New System.Net.NetworkCredential(user_ftp, passw_ftp)
    '        conect_ftp.UseBinary = True


    '        Dim fichero As Byte() = System.IO.File.ReadAllBytes(imagen_up)
    '        Dim requeststream As System.IO.Stream = conect_ftp.GetRequestStream

    '        For offset As Integer = 0 To fichero.Length Step 1024
    '            BackgroundWorker1.ReportProgress(CType(offset * progreso_img_pricipal.Maximum / fichero.Length, Integer))

    '            Dim chsize As Integer = fichero.Length - offset
    '            If chsize > 1024 Then chsize = 1024
    '            requeststream.Write(fichero, offset, chsize)


    '        Next
    '        requeststream.Close()
    '        requeststream.Dispose()


    '    Catch ex As Exception
    '        MsgBox("error al subir la imagen : " & "/img/p" & path & "/" & idimage & ".jpg" & "/br" & ex.Message)

    '    End Try
    'End Sub

    
   
    
    


End Class