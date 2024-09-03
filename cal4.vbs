class a
    private m_name
    private m_sex

    public property get Name
        Name = m_name
    end property
    public property let Name(value)
          m_name = value
    end property
end class

dim x 

set x = new a
x.Name = "hellp"
wscript.echo x.Name

set x= nothing
